VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRATurnovrRt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Turnover Rates Report - Active Employee"
   ClientHeight    =   9000
   ClientLeft      =   570
   ClientTop       =   1095
   ClientWidth     =   10425
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
   ScaleHeight     =   9000
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   52
      Top             =   8700
      Width           =   10095
   End
   Begin VB.VScrollBar scrControl 
      Height          =   8115
      LargeChange     =   315
      Left            =   9840
      Max             =   100
      SmallChange     =   315
      TabIndex        =   51
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
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
         Top             =   8130
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
         Top             =   7815
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2295
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
         Top             =   6720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame frmTerm 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   3420
         TabIndex        =   3
         Top             =   6690
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
         Left            =   4830
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
         ItemData        =   "fzATrnRt.frx":0000
         Left            =   2280
         List            =   "fzATrnRt.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Select Date Range Based On"
         Top             =   4290
         Width           =   2325
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1980
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
         Left            =   1980
         TabIndex        =   12
         Tag             =   "00-Enter Status Code"
         Top             =   5340
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   1980
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
         Left            =   3660
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
         Left            =   1980
         TabIndex        =   24
         Tag             =   "01-Termination Code"
         Top             =   4860
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
         Index           =   9
         Left            =   1980
         TabIndex        =   25
         Tag             =   "00-Enter Status Code"
         Top             =   6180
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   26
         Tag             =   "Detailed Report"
         Top             =   7140
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
         Left            =   3495
         TabIndex        =   27
         Tag             =   "Summary Report"
         Top             =   7140
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   1980
         TabIndex        =   28
         Tag             =   "00-Enter Status Code"
         Top             =   5760
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   7605
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
         TabIndex        =   44
         Top             =   7845
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
         TabIndex        =   43
         Top             =   8160
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   4320
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   4905
         Width           =   1500
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Time Status"
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
         Top             =   5385
         Width           =   1125
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
         TabIndex        =   31
         Top             =   6225
         Width           =   885
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
         Left            =   180
         TabIndex        =   30
         Top             =   7140
         Width           =   1065
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Time Status"
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
         Top             =   5805
         Width           =   1170
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9840
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
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRATurnovrRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cnRun  As ADODB.Connection
Attribute cnRun.VB_VarHelpID = -1
Dim glbstrSelCri_1 As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Emergency Contact Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
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
MsgBox "Error Printing - check your Windows Printer setup"
Resume Next

End Sub

Public Sub cmdView_Click()
Dim X%, selected&
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
    X% = Cri_SetAll()
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
Dim X%
Dim vPosGroup As String
    
    'Hemu 06/02/2004 Begin
    cmbDateBased.AddItem lStr("Original Hire Date")
    cmbDateBased.AddItem lStr("Seniority Date")
    cmbDateBased.AddItem lStr("Last Hire Date")
    cmbDateBased.AddItem lStr("Last Day Date")
    cmbDateBased.AddItem lStr("Union Date")
    cmbDateBased.AddItem lStr("User Defined Date")
    If glbCompSerial <> "S/N - 2347W" And glbCompSerial <> "S/N - 2394W" Then
        cmbDateBased.AddItem lStr("Termination Date")
    End If
    
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
    'CodeCri = "({" & strCd$ & "} = '" & clpCode(intIdx%).Text & "')"
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
'glbstrSelCri_1 = glbstrSelCri

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

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

Private Sub Cri_FullPartTime()
Dim EECri As String, OneSet%, X%

    If Len(clpCode(2).Text) < 1 Then GoTo Checkother

    If glbOracle Then
        If Len(clpCode(8).Text) < 1 Then
            If Len(clpCode(9).Text) < 1 Then
                EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(2).Text) & "']"
            Else
                EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(2).Text) & "'" & ",'" & getCodes(clpCode(8).Text) & "','" & getCodes(clpCode(9).Text) & "']"
            End If
        Else
            If Len(clpCode(9).Text) < 1 Then
                EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(2).Text) & "'" & ",'" & getCodes(clpCode(8).Text) & "']"
            Else
                EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(2).Text) & "'" & ",'" & getCodes(clpCode(8).Text) & "','" & getCodes(clpCode(9).Text) & "']"
            End If
        End If
    Else
        If Len(clpCode(8).Text) < 1 Then
            If Len(clpCode(9).Text) < 1 Then
                EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(2).Text) & "')"
            Else
                EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(2).Text) & "','" & getCodes(clpCode(9).Text) & "')"
            End If
        Else
            If Len(clpCode(9).Text) < 1 Then
                EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(2).Text) & "'" & ",'" & getCodes(clpCode(8).Text) & "')"
            Else
                EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(2).Text) & "'" & ",'" & getCodes(clpCode(8).Text) & "','" & getCodes(clpCode(9).Text) & "')"
            End If
        End If
    End If
    GoTo set_query
    
Checkother:
    If Len(clpCode(8).Text) < 1 Then GoTo CheckOther1

    'EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
    If glbOracle Then
        If Len(clpCode(9).Text) < 1 Then
            EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(8).Text) & "']"
        Else
            EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(8).Text) & "'" & ",'" & getCodes(clpCode(9).Text) & "']"
        End If
    Else
        If Len(clpCode(9).Text) < 1 Then
            EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(8).Text) & "')"
        Else
            EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(8).Text) & "'" & ",'" & getCodes(clpCode(9).Text) & "')"
        End If
    End If
    GoTo set_query
    
    
CheckOther1:
    If Len(clpCode(9).Text) < 1 Then Exit Sub
    
    If glbOracle Then
        If Len(clpCode(9).Text) < 1 Then
            EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(9).Text) & "']"
        Else
            EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(9).Text) & "'" & ",'" & getCodes(clpCode(2).Text) & "']"
        End If
    Else
        If Len(clpCode(9).Text) < 1 Then
            EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(9).Text) & "')"
        Else
            EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(9).Text) & "'" & ",'" & getCodes(clpCode(2).Text) & "')"
        End If
    End If

set_query:
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
    glbstrSelCri_1 = glbstrSelCri
End Sub

Private Sub Cri_OtherCode()
Dim EECri As String, OneSet%, X%

    If Len(clpCode(9).Text) < 1 Then Exit Sub

    'Hemu 06/02/2004 Begin
    'EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
    If glbOracle Then
        EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpCode(9).Text) & "']"
    Else
        EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpCode(9).Text) & "')"
    End If
    'Hemu 06/02/2004 End
    
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True

End Sub

Private Sub Cri_TermCode()
Dim EECri As String, OneSet%, X%

    If Len(clpCode(7).Text) < 1 Then Exit Sub

    'Hemu 06/02/2004 Begin
    'EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
    If glbOracle Then
        EECri = "{HREMP.ED_EMP} IN ['" & getCodes(clpCode(7).Text) & "']"
    Else
        EECri = "{HREMP.ED_EMP} IN ('" & getCodes(clpCode(7).Text) & "')"
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
Dim X%, strRName$, selected&
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
For X% = 0 To 4
    If X% <> 2 Then
        Call Cri_Code(X%)
    End If
Next X%
Call Cri_Code(6)
Call Cri_PT
Call Cri_Shift
Call Cri_EE
'Call Cri_TermCode
Call Cri_FullPartTime
''Call Cri_OtherCode

'Hemu 06/03/2004 Begin
'As of Date = Date Range
If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    Select Case cmbDateBased
    Case lStr("Original Hire Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Seniority Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Last Hire Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Union Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("User Defined Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Last Day Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LDAY} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {Term_HREMP.ED_LDAY} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LDAY} >= " & Date_SQL(dlpDateRange(2)) & " AND {Term_HREMP.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {Term_HREMP.ED_LDAY} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    End Select
    
    If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2394W" Then
        glbstrSelCri = Replace(glbstrSelCri, "Term_HREMP.", "")
    End If
    
End If
'Hemu 06/03/2004 End

Call SETWRK

' report name
If comGroup(0) <> "(none)" Then
    If optGrouping(0) Then
        If glbCompSerial = "S/N - 2394W" Then 'Ticket #15646
            strRName$ = glbIHRREPORTS & "rzTrnoR1_S.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzTrnoR1_D.rpt"
        End If
    Else
        strRName$ = glbIHRREPORTS & "rzTrnoR1_1.rpt"
    End If
Else
    If optGrouping(0) Then
        strRName$ = glbIHRREPORTS & "rzTrnoRt_D.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzTrnoRt_0.rpt"
    End If
End If
Me.vbxCrystal.ReportFileName = strRName$

' set to sorting/grouping criteria
X% = Cri_Sorts()

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    'If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
    '    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_WRKEMP}='" & glbUserID & "'"
    'Else
        Me.vbxCrystal.SelectionFormula = "{HREMP.ED_WRKEMP}='" & glbUserID & "'"
    'End If
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
        Me.vbxCrystal.DataFiles(11) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(12) = glbIHRAUDIT
    End If
    ' window title if appropriate
    Me.vbxCrystal.WindowTitle = "Turnover Rates Reports"
    
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

ESQLQ = glbstrSelCri_1
ESQLQ = Replace(ESQLQ, "{", "")
ESQLQ = Replace(ESQLQ, "}", "")
ESQLQ = Replace(ESQLQ, "Term_HREMP.", "")
ESQLQ = Replace(ESQLQ, "HREMP.", "")
cnRun.BeginTrans
cnRun.Execute "DELETE FROM HREMP_HS WHERE ED_WRKEMP='" & glbUserID & "'"
cnRun.CommitTrans

MDIMain.panHelp(0).FloodPercent = 30

'for active employees
If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2394W" Then
    SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
    SQLQ = SQLQ & " SELECT " & xFieldList
    SQLQ = SQLQ & ",'1_'  AS KEY_EMPNBR "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
    SQLQ = SQLQ & " FROM HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & "WHERE" & ESQLQ
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
    
    'rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS", cnRun, adOpenStatic, adLockPessimistic
    If glbOracle Then
        rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS WHERE SUBSTR(KEY_EMPNBR,1,1)='0'", cnRun, adOpenStatic, adLockPessimistic
    Else
        rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS WHERE LEFT(KEY_EMPNBR,1)='0'", cnRun, adOpenStatic, adLockPessimistic
    End If

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
Else
MDIMain.panHelp(0).FloodPercent = 50

'for terminated employees
'If chkTerm Then    'Hemu
    SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
    SQLQ = SQLQ & "SELECT " & xFieldList
    SQLQ = SQLQ & ",'0_'  AS KEY_EMPNBR "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
    SQLQ = SQLQ & " FROM Term_HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "WHERE" & ESQLQ
    
    Select Case cmbDateBased
    Case lStr("Termination Date")
        If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(2))
            SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(3)) & ")"
        Else
            If IsDate(dlpDateRange(2)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(2)) & ")"
            End If
            If IsDate(dlpDateRange(3)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(3)) & ")"
            End If
        End If
    End Select
    
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
End If 'Hemu

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

End Sub

Private Function CriCheck()
Dim X%

CriCheck = False

If Not clpDiv.ListChecker Then
    Exit Function
End If

If Not clpDept.ListChecker Then
    Exit Function
End If

For X% = 0 To 6
    If Not clpCode(X%).ListChecker Then
        Exit Function
    End If
Next X%

If Not clpPT.ListChecker Then
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not clpCode(7).ListChecker Then
    Exit Function
End If

If Len(Trim(dlpDateRange(2).Text)) = 0 Then
    MsgBox "From Date cannot be blank"
    dlpDateRange(2).SetFocus
    Exit Function
End If
If Len(Trim(dlpDateRange(3).Text)) = 0 Then
    MsgBox "To Date cannot be blank"
    dlpDateRange(3).SetFocus
    Exit Function
End If

If Len(Trim(clpCode(7).Text)) = 0 Then
    MsgBox "Voluntary Separation Termination Reason cannot be blank"
    clpCode(7).SetFocus
    Exit Function
End If

'If Not clpCode(8).ListChecker Then
'    Exit Function
'End If
'If Len(Trim(clpCode(8).Text)) = 0 Then
'    MsgBox "Involuntary Separation Termination Reason cannot be blank"
'    clpCode(8).SetFocus
'    Exit Function
'End If

If Not clpCode(2).ListChecker Then
    Exit Function
End If
If Not clpCode(8).ListChecker Then
    Exit Function
End If
If Not clpCode(9).ListChecker Then
    Exit Function
End If

If Len(Trim(clpCode(2).Text)) = 0 And Len(Trim(clpCode(8).Text)) = 0 And Len(Trim(clpCode(9).Text)) = 0 Then
    MsgBox "All Full Time, Part Time and Other Status cannot be blank"
    clpCode(2).SetFocus
    Exit Function
End If

'If Len(Trim(clpCode(9).Text)) = 0 Then
'    MsgBox "Other Status cannot be blank"
'    clpCode(9).SetFocus
'    Exit Function
'End If

CriCheck = True
End Function

Private Sub dlpDateRange_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
glbOnTop = "FRMRATURNOVRRT"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False


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
scrFrame.Height = 8535
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8400 Then
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
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 9750 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 150
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
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%, dscGroup1
Dim SubTotal
'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
' first set primary grouping

X% = 0
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
        Case lStr("Termination Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Termination Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Termination Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Termination Date") & ")'"
            End If
        Case lStr("User Defined Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("User Defined Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("User Defined Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("User Defined Date") & ")'"
            End If
        Case lStr("Last Day Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Last Day Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Last Day Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Last Day Date") & ")'"
            End If
    End Select
End If

Me.vbxCrystal.Formulas(10) = "lblLDay='" & lStr("Last Day") & "'"
Me.vbxCrystal.Formulas(11) = "lblPT='" & lStr("Category") & "'"
Me.vbxCrystal.Formulas(12) = "FromDate='" & dlpDateRange(2).Text & "'"
Me.vbxCrystal.Formulas(13) = "ToDate='" & dlpDateRange(3).Text & "'"

Me.vbxCrystal.Formulas(2) = "Title='Turnover Rates Report'"
If Len(clpCode(2).Text) < 1 Then
    If Len(clpCode(8).Text) < 1 Then
        Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(9).Text & "'"
    Else
        If Len(clpCode(9).Text) < 1 Then
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(8).Text & "'"
        Else
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(8).Text & "," & clpCode(9).Text & "'"
        End If
    End If
Else
    If Len(clpCode(8).Text) < 1 Then
        If Len(clpCode(9).Text) < 1 Then
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(2).Text & "'"
        Else
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(2).Text & "," & clpCode(9).Text & "'"
        End If
    Else
        If Len(clpCode(9).Text) < 1 Then
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(2).Text & "," & clpCode(8).Text & "'"
        Else
            Me.vbxCrystal.Formulas(9) = "selectcriteria='Category Selected : " & clpCode(2).Text & "," & clpCode(8).Text & "," & clpCode(9).Text & "'"
        End If
    End If
End If

If comGroup(0) = "(none)" Then
    GoTo Set_Criteria
    Exit Function
End If

Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

'Total Employee - Count
dscGroup$ = "Count ({HREMP.ED_EMPNBR}, " & grpField$ & ")"
dscGroup$ = "G1TotalEmp=" & dscGroup$
Me.vbxCrystal.Formulas(3) = dscGroup$

Set_Criteria:
'Voluntary - FT/PT Status Codes:
'If Len(Trim(clpCode(2).Text)) > 0 Then
'    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
    
        dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_EMP} = ")
        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {Term_HRTRMEMP.Term_Reason} = ")
        dscGroup$ = "TotFTPT=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & dscGroup1 & "') then 1 else 0"
        
        If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2394W" Then
            'Full Time
            dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_PT} = ")
            dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HREMP.ED_EMP} = ")
            'dscGroup$ = "TotFTPT=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "') then 1 else 0"
            dscGroup$ = "TotFTPT=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "')"
            
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                'dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP_1.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} in DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")" & " to DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} <= DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} >= DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")"
            End If
            dscGroup$ = dscGroup$ & " then 1 else 0"
            
            Me.vbxCrystal.Formulas(14) = dscGroup$
            
            'Part Time
            dscGroup$ = Replace(getCodes(clpCode(8).Text), ",", " OR {HREMP.ED_PT} = ")
            dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HREMP.ED_EMP} = ")
            'dscGroup$ = "TotFTPT=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "') then 1 else 0"
            dscGroup$ = "TotPT=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "')"
            
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                'dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP_1.ED_LDAY} <= " & Date_SQL(dlpDateRange(3))
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} in DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")" & " to DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} <= DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} >= DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")"
            End If
            dscGroup$ = dscGroup$ & " then 1 else 0"
        End If
'    Else
'        dscGroup$ = "TotFTPT=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(2).Text) & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & getCodes(clpCode(7).Text) & "') then 1 else 0"
'    End If
    Me.vbxCrystal.Formulas(5) = dscGroup$
'End If

'Voluntary - NOT FT/PT Status Codes:
'If Len(Trim(clpCode(2).Text)) > 0 Then
'    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
        dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_EMP} = ")
        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {Term_HRTRMEMP.Term_Reason} = ")
        dscGroup$ = "TotOthStatus=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & dscGroup1 & "') then 1 else 0"
        
        If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2394W" Then
            dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_PT} = ")
            dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HREMP.ED_EMP} = ")
            'dscGroup$ = "TotOthStatus=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "') then 1 else 0"
            dscGroup$ = "TotOthStatus=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "')"
            
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} in DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")" & " to DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} <= DateTime(" & Year(dlpDateRange(3)) & "," & month(dlpDateRange(3)) & "," & Day(dlpDateRange(3)) & ")"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                dscGroup$ = dscGroup$ & " AND {HREMP_1.ED_LDAY} >= DateTime(" & Year(dlpDateRange(2)) & "," & month(dlpDateRange(2)) & "," & Day(dlpDateRange(2)) & ")"
            End If
            dscGroup$ = dscGroup$ & " then 1 else 0"
            
        End If
'    Else
'        dscGroup$ = "TotOthStatus=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(9).Text) & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & getCodes(clpCode(7).Text) & "') then 1 else 0"
'    End If
    Me.vbxCrystal.Formulas(6) = dscGroup$
'End If


''Involuntary FT/PT Status Codes:
''If Len(Trim(clpCode(2).Text)) > 0 Then
''    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
'        dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_EMP} = ")
'        dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {Term_HRTRMEMP.Term_Reason} = ")
'        dscGroup$ = "TotFTPT1=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & dscGroup1 & "') then 1 else 0"
'
'        If glbCompSerial = "S/N - 2347W" Then
'            dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_PT} = ")
'            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HREMP.ED_EMP} = ")
'            dscGroup$ = "TotFTPT1=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "') then 1 else 0"
'        End If
''    Else
''        dscGroup$ = "TotFTPT1=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(2).Text) & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & getCodes(clpCode(8).Text) & "') then 1 else 0"
''    End If
'    Me.vbxCrystal.Formulas(7) = dscGroup$
''End If
'
''Involuntary NOT FT/PT Status Codes:
''If Len(Trim(clpCode(2).Text)) > 0 Then
''    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
'        dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_EMP} = ")
'        dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {Term_HRTRMEMP.Term_Reason} = ")
'        dscGroup$ = "TotOthStatus1=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & dscGroup1 & "') then 1 else 0"
'
'        If glbCompSerial = "S/N - 2347W" Then
'            dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_PT} = ")
'            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HREMP.ED_EMP} = ")
'            dscGroup$ = "TotOthStatus1=if ({HREMP.ED_PT} = '" & dscGroup$ & "') AND ({HREMP.ED_EMP} = '" & dscGroup1 & "') then 1 else 0"
'        End If
''    Else
''        dscGroup$ = "TotOthStatus1=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(9).Text) & "') AND ({Term_HRTRMEMP.Term_Reason} = '" & getCodes(clpCode(8).Text) & "') then 1 else 0"
''    End If
'    Me.vbxCrystal.Formulas(8) = dscGroup$
''End If

'Total Separation
If comGroup(0) = "(none)" Then
    'dscGroup$ = "Sum({@TotFTPT}) + Sum({@TotOthStatus}) + Sum({@TotFTPT1}) + Sum({@TotOthStatus1})"
    dscGroup$ = "Sum({@TotFTPT}) + Sum({@TotPT}) + Sum({@TotOthStatus})"
    dscGroup$ = "TotSeparation=" & dscGroup$
Else
    'dscGroup$ = "Sum({@TotFTPT}, " & grpField$ & ") + Sum({@TotOthStatus}, " & grpField$ & ") + Sum({@TotFTPT1}, " & grpField$ & ") + Sum({@TotOthStatus1}, " & grpField$ & ")"
    dscGroup$ = "Sum({@TotFTPT}, " & grpField$ & ") + Sum({@TotPT}, " & grpField$ & ") + Sum({@TotOthStatus}, " & grpField$ & ")"
    dscGroup$ = "G1TotSeparation=" & dscGroup$
End If
Me.vbxCrystal.Formulas(4) = dscGroup$

Cri_Sorts = z% ' next section number to format

End Function


