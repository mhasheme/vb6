VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRBenefits 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Benefits Report"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   1005
   ClientWidth     =   12255
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
   ScaleHeight     =   10755
   ScaleWidth      =   12255
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   60
      Top             =   9855
      Width           =   9855
   End
   Begin VB.VScrollBar scrControl 
      Height          =   9555
      LargeChange     =   315
      Left            =   9840
      Max             =   100
      SmallChange     =   315
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   9615
      Begin VB.CheckBox chkEndedBenefit 
         Caption         =   "Exclude Ended Benefits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2120
         TabIndex        =   23
         Top             =   6600
         Width           =   2055
      End
      Begin VB.TextBox txtTaxPerc 
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
         Left            =   2120
         MaxLength       =   7
         TabIndex        =   16
         Tag             =   "10-Enter Tax Percentage"
         Top             =   4946
         Visible         =   0   'False
         Width           =   870
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "First Level of grouping records"
         Top             =   8280
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "Second level of grouping records"
         Top             =   8595
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
         Index           =   2
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Tag             =   "Third level of grouping records"
         Top             =   8910
         Width           =   2325
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
         Index           =   3
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "Third level of grouping records"
         Top             =   9225
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5480
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "00-Employee Position Shift"
         Top             =   4612
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtCoverageType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2120
         MaxLength       =   1
         TabIndex        =   14
         Tag             =   "00-Type of Coverage (S, F, W, X, etc)"
         Top             =   4612
         Width           =   330
      End
      Begin VB.CheckBox chkDOB 
         Caption         =   "Show Birth Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   6600
         Width           =   2055
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   225
         Left            =   2115
         TabIndex        =   35
         Top             =   7020
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   397
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optAnnMth 
            Height          =   255
            Index           =   1
            Left            =   1485
            TabIndex        =   25
            Tag             =   "Show Dollars in Annual"
            Top             =   0
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Annual"
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
         Begin Threed.SSOption optAnnMth 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Tag             =   "Show Dollars in Monthly"
            Top             =   0
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Monthly"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Tag             =   "00-Specific Benefit Desired"
         Top             =   1940
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "BNCD"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   4
         Tag             =   "00-Status Codes"
         Top             =   1606
         Width           =   7215
         _ExtentX        =   12726
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
         Left            =   1800
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1272
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   27
         Tag             =   "Summary Report"
         Top             =   7380
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   2
         Left            =   5070
         TabIndex        =   28
         Tag             =   "Beneficiary Report"
         Top             =   7380
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Beneficiary"
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
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   604
         Width           =   7215
         _ExtentX        =   12726
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
         Left            =   1800
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   270
         Width           =   7215
         _ExtentX        =   12726
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
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Tag             =   "EDLC-Location"
         Top             =   938
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Tag             =   "EDPT-Category"
         Top             =   2274
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   12
         Tag             =   "EDSE-Section "
         Top             =   3944
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   1800
         TabIndex        =   11
         Tag             =   "EDAB-Administered By"
         Top             =   3610
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   10
         Tag             =   "EDRG-Region"
         Top             =   3276
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   9
         Tag             =   "40-Date upto and including this date forward"
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Tag             =   "40-Date from and including this date forward"
         Top             =   2942
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1135
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Tag             =   "10-Enter Employee Number"
         Top             =   2608
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   6895
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   13
         Tag             =   "00-Enter Benefit Group Code"
         Top             =   4278
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "BGMF"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   26
         Tag             =   "Detailed Report"
         Top             =   7380
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
      End
      Begin INFOHR_Controls.CodeLookup clpGLNum 
         Height          =   285
         Left            =   5160
         TabIndex        =   17
         Tag             =   "00-General Ledger - Code"
         Top             =   4950
         Visible         =   0   'False
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   18
         Tag             =   "41-Benefit End Date"
         Top             =   5280
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1300
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   19
         Tag             =   "41-Benefit End Date"
         Top             =   5280
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1300
      End
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Tag             =   "31-Province of Residence - Code"
         Top             =   5610
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin INFOHR_Controls.CodeLookup clpProvEmp 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Tag             =   "31-Province of Employment - Code"
         Top             =   5940
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   29
         Tag             =   "Beneficiary Report"
         Top             =   7380
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Billing Statement"
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
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   65
         Top             =   5985
         Width           =   1455
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
         Left            =   120
         TabIndex        =   64
         Top             =   5655
         Width           =   1365
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   5325
         Width           =   675
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G/L #"
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
         Left            =   4080
         TabIndex        =   62
         Top             =   4995
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Percentage"
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
         Top             =   4991
         Visible         =   0   'False
         Width           =   1140
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
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   8020
         Width           =   1575
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
         Left            =   135
         TabIndex        =   56
         Top             =   315
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
         Left            =   135
         TabIndex        =   55
         Top             =   649
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   135
         TabIndex        =   54
         Top             =   1317
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   3
         Left            =   135
         TabIndex        =   53
         Top             =   1651
         Width           =   450
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Code"
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
         Index           =   4
         Left            =   135
         TabIndex        =   52
         Top             =   1985
         Width           =   915
      End
      Begin VB.Label lblTitle 
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
         Index           =   6
         Left            =   135
         TabIndex        =   51
         Top             =   2653
         Width           =   1290
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From  / To Date"
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
         Index           =   7
         Left            =   135
         TabIndex        =   50
         Top             =   2987
         Width           =   1140
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   7410
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   135
         TabIndex        =   48
         Top             =   8310
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #2"
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
         Index           =   10
         Left            =   135
         TabIndex        =   47
         Top             =   8625
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #3"
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
         Index           =   11
         Left            =   135
         TabIndex        =   46
         Top             =   8940
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   135
         TabIndex        =   45
         Top             =   9255
         Width           =   825
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   135
         TabIndex        =   44
         Top             =   983
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
         Left            =   135
         TabIndex        =   43
         Top             =   3321
         Width           =   510
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
         Left            =   135
         TabIndex        =   42
         Top             =   3989
         Width           =   540
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
         Left            =   135
         TabIndex        =   41
         Top             =   3655
         Width           =   1125
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
         Left            =   135
         TabIndex        =   40
         Top             =   2319
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
         Left            =   4080
         TabIndex        =   39
         Top             =   4650
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblBGOld 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Group"
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
         Top             =   4323
         Width           =   975
      End
      Begin VB.Label lblCoverage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Show Dollars in"
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
         Top             =   7020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage Type"
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
         Top             =   4657
         Width           =   1185
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10200
      Top             =   7560
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
   Begin Crystal.CrystalReport vbxCrystal1 
      Left            =   10200
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc, xKeyPress, xTxtDiv, xLblDivDesc  'js-20Apr99
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName
Dim flgNiagara As Boolean

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Benefits Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    
    If glbCompSerial = "S/N - 2362W" Then  'City of Sarnia Ticket #23890 Franks 07/08/2013
        If optGrouping(3).Value Then
            MDIMain.Timer1.Enabled = True
            Call set_PrintState(True)
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    'City of Niagara Falls
    If flgNiagara And glbCompSerial = "S/N - 2276W" Then
        Me.vbxCrystal1.Destination = 1
        Me.vbxCrystal1.Action = 1
        vbxCrystal1.Reset
        flgNiagara = False
    End If
    
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
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
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    X% = Cri_SetAll()
    
    If glbCompSerial = "S/N - 2362W" Then  'City of Sarnia Ticket #23890 Franks 07/08/2013
        If optGrouping(3).Value Then
            MDIMain.Timer1.Enabled = True
            Call set_PrintState(True)
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    
    Me.vbxCrystal.Destination = 0

    MDIMain.Timer1.Enabled = False
    
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    'City of Niagara Falls
    If flgNiagara And glbCompSerial = "S/N - 2276W" Then
        Me.vbxCrystal1.Destination = 0
        Me.vbxCrystal1.Action = 1
        vbxCrystal1.Reset
        flgNiagara = False
    End If
    
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
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim X%

    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(3).Clear
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14189
        If optGrouping(1).Value Then
            comGroup(0).AddItem lStr("G/L #") 'Ticket #25466 Franks 05/09/2014
        End If
    End If
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Benefit Code"
    comGroup(1).AddItem "Benefit Group"
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    comGroup(2).AddItem "Benefit Code"
    comGroup(2).AddItem "(none)"
    comGroup(2).ListIndex = 0
    comGroup(3).AddItem "Effective Date"
    comGroup(3).ListIndex = 0

End Sub

Private Sub Cri_BeneGroup()
Dim EECri As String, OneSet%, X%
Dim strCd$

X% = 8
If Len(clpCode(X%).Text) = 0 Then Exit Sub

If Len(clpCode(X%).Text) > 0 Then
    EECri = EECri & "'" & Replace(clpCode(X%).Text, ",", "','") & "'"
End If
If Right(EECri, 1) = "," Then EECri = Left(EECri, Len(EECri) - 1)


If optGrouping(2) Then
    strCd$ = "HREMP.ED_BENEFIT_GROUP"
Else
    strCd$ = "HRBENFT.BF_GROUP"
End If

EECri = "({" & strCd$ & "} in ['" & Replace(clpCode(X%).Text, ",", "','") & "'])"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Benefit()
Dim EECri As String, OneSet%, X%


If Len(clpCode(1).Text) = 0 Then Exit Sub

For X% = 1 To 1
    If Len(clpCode(X%).Text) > 0 Then
        If glbLinamar Then
            If InStr(clpCode(1), ",") <> 0 Then
                EECri = EECri & "'" & clpDiv.Text & Replace(clpCode(X%).Text, ",", "','" & clpDiv.Text) & "'"
            Else
                EECri = EECri & "'" & clpDiv.Text & clpCode(X%).Text & "',"
            End If
        Else
            EECri = EECri & "'" & Replace(clpCode(X%).Text, ",", "','") & "'"
        End If
    End If
    If Right(EECri, 1) = "," Then EECri = Left(EECri, Len(EECri) - 1)
Next X%

If optGrouping(2) Then
  EECri = "{HRBENS.BD_BCODE} IN [" & EECri & "]"
Else
  EECri = "{HRBENFT.BF_BCODE} IN [" & EECri & "]"
End If

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
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_ORG"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 7 Then strCd$ = "HREMP.ED_SECTION"
    If intIdx% = 9 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 10 Then strCd$ = "HREMP.ED_ADMINBY"
    CodeCri = "({" & strCd$ & "} in ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
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

Private Sub Cri_GLNO()
Dim GLNOCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpGLNum.Text) > 0 Then
    GLNOCri = "({HREMP.ED_GLNO} = '" & clpGLNum.Text & "')"
End If

If Len(GLNOCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = GLNOCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & GLNOCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
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

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HRBENFT.BF_EDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDateRange(X%).Text) > 0 Then
        TempCri = "({HRBENFT.BF_EDATE}  "
        If X% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(X%).Text)
        dtMM% = month(dlpDateRange(X%).Text)
        dtDD% = Day(dlpDateRange(X%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next X%

Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Benefit_End_Date()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDate(0).Text) > 0 And Len(dlpDate(1).Text) > 0 Then
    TempCri = "({HRBENFT.BF_CEASEDATE} "
    dtYYY% = Year(dlpDate(0).Text)
    dtMM% = month(dlpDate(0).Text)
    dtDD% = Day(dlpDate(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDate(1).Text)
    dtMM% = month(dlpDate(1).Text)
    dtDD% = Day(dlpDate(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDate(X%).Text) > 0 Then
        TempCri = "({HRBENFT.BF_CEASEDATE}  "
        If X% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDate(X%).Text)
        dtMM% = month(dlpDate(X%).Text)
        dtDD% = Day(dlpDate(X%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next X%


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Excluded_Ended_Benefit()
Dim EECri As String
Dim dtYYY%, dtMM%, dtDD%
    If chkEndedBenefit.Value = 1 Then
        EECri = "(ISNULL({HRBENFT.BF_CEASEDATE}) OR {HRBENFT.BF_CEASEDATE} > "
        
        dtYYY% = Year(Now)
        dtMM% = month(Now)
        dtDD% = Day(Now)
        EECri = EECri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        
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

Private Function Cri_SetAll()
Dim X%
Dim LocstrSelCri 'Ticket #27375 Franks 08/06/2015

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
flgNiagara = False

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Code(0)
Call Cri_Code(5)
Call Cri_Code(6)
Call Cri_Code(7)
'Hemu
'Ticket #13187 Frank 06/15/07, multi Benefit Group Code select
'Call Cri_Code(8)
Call Cri_BeneGroup
'Hemu
Call Cri_Code(9)
Call Cri_Code(10)
Call Cri_EE
Call Cri_Benefit
If Not optGrouping(2) Then Call Cri_FTDates 'changed by RAUBREY 5/14/97
Call Cri_PT
Call Cri_Shift
Call Cri_GLNO 'Ticket #14189

If Not optGrouping(2) Then Call Cri_CoverageType

Call Cri_Benefit_End_Date
Call Cri_Excluded_Ended_Benefit

Call Cri_ProvResidence
Call Cri_ProvEmployment

If glbCompSerial = "S/N - 2362W" Then  'City of Sarnia Ticket #23890 Franks 07/08/2013
    If optGrouping(3).Value Then
        Call CitySarniaBillingWrk
        
        'Ticket #24106 - Add Terminated Employees
        Call CitySarniaBillingWrk_TerminatedEmployees
        
        Call CitySarniaGenerate_Excel_File
        Exit Function
    End If
End If

If optAnnMth(1).Value = True Then
    ' report name
    If optGrouping(0) Then
        Me.vbxCrystal.Formulas(10) = "showdob= " & IIf(chkDOB, "True", "False") & "" 'Ticket #13187
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftd1.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftdt.rpt"
        End If
    ElseIf optGrouping(1) Then
        Me.vbxCrystal.Formulas(10) = "showdob= " & IIf(chkDOB, "True", "False") & "" 'Ticket #13187
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbenef1.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbeneft.rpt"
        End If
    Else
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftb1.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftdb.rpt" 'beneficiary report added by RAUBREY 5/14/97
        End If
    End If
ElseIf optAnnMth(0).Value = True Then
    'For Monthly reports
    If optGrouping(0) Then
        Me.vbxCrystal.Formulas(10) = "showdob= " & IIf(chkDOB, "True", "False") & "" 'Ticket #13187
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbftd1M.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbftdtM.rpt"
        End If
    ElseIf optGrouping(1) Then
        Me.vbxCrystal.Formulas(10) = "showdob= " & IIf(chkDOB, "True", "False") & "" 'Ticket #13187
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbenf1M.rpt"
            'City of Niagara Falls
            If glbCompSerial = "S/N - 2276W" Then
                Me.vbxCrystal1.Formulas(10) = "showdob= " & IIf(chkDOB, "True", "False") & "" 'Ticket #13187
                Me.vbxCrystal1.ReportFileName = glbIHRREPORTS & "SN2276_rzbenf1M.rpt"
                Me.vbxCrystal1.SelectionFormula = glbstrSelCri
                Me.vbxCrystal1.Connect = RptODBC_SQL
                Me.vbxCrystal1.WindowTitle = "Benefits Report"
                flgNiagara = True
            End If
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbenftM.rpt"
        End If
    Else
        If comGroup(0).Text = "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftb1.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzbnftdb.rpt" 'beneficiary report added by RAUBREY 5/14/97
        End If
    End If
End If

If Len(Trim(txtTaxPerc)) > 0 Then
    If optGrouping(1) And IsNumeric(txtTaxPerc) Then
        Me.vbxCrystal.Formulas(11) = "comptaxperc= " & txtTaxPerc.Text & ""  'Ticket #14065
        Me.vbxCrystal.Formulas(12) = "emptaxperc= " & txtTaxPerc.Text & ""  'Ticket #14065
    End If
End If
' set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

'Ticket #27375 Franks 08/06/2015 - begin
If glbWFC Then
    If glbNoNONE And glbNoEXEC Then     ' -EXE
        LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"    'Hemu -EXE
    ElseIf glbNoNONE Then
        LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
    ElseIf glbNoEXEC Then    'Hemu -EXE
        LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"    'Hemu -EXE
    Else
        LocstrSelCri = glbstrSelCri
    End If
    glbstrSelCri = LocstrSelCri
End If
'Ticket #27375 Franks 08/06/2015 - end

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

' set security for database

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Benefits Report"

Cri_SetAll = True


Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
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

Private Sub Cri_CoverageType()
Dim EECri As String, OneSet%, X%

If Len(txtCoverageType.Text) < 1 Then Exit Sub
EECri = "{HRBENFT.BF_COVER}= '" & txtCoverageType.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
Dim grpField01$

'for labels - sort by name always
' imbeded in report
Cri_Sorts = 0
' first set primary grouping
grpField$ = getEGroup(comGroup(0).Text)
grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")

If comGroup(0) <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    
    'Ticket #14485 - Jerry wants GL # and Description to be part of Group Header
    'This does not work on run time so had to change the default report grouping in the
    'Crystal Report file to GL # and Description. So when GL# is choosen, just pass the
    'label and let the default report option display the grouping of GL # and Description.
    If glbCompSerial = "S/N - 2388W" And optGrouping(1).Value And comGroup(0).Text = lStr("G/L #") Then  'DNSSAB Ticket #14189 'Ticket #25466 Franks 05/09/2014
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & " - Description : '"
        Me.vbxCrystal.Formulas(0) = dscGroup$
    Else
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal.Formulas(0) = dscGroup$
        
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    End If
    
    strSFormat$ = "GH1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = strSFormat$
    
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = strSFormat$
    
    
    'Ticket #22682 - Release 8.0 - Show Grouping 1 totals
    If InStr(1, Me.vbxCrystal.ReportFileName, "rzbenftM.rpt") > 0 Or InStr(1, Me.vbxCrystal.ReportFileName, "rzbeneft.rpt") > 0 Then
        If grpField$ = "{HRGL.GL_DESCR}" Then 'Ticket #25466 Franks 05/23/2014
            grpField$ = "{HRGL.GL_NO}"
        End If
        Me.vbxCrystal.Formulas(3) = "group1total=Sum ({HRBENFT.BF_CCOST}, " & grpField$ & ")+Sum ({HRBENFT.BF_ECOST}, " & grpField$ & ")"
        
        Me.vbxCrystal.Formulas(4) = "group1_emptax=Sum ({HRBENFT.BF_ECOST}, " & grpField$ & ")*({@emptaxperc}/100)"
        Me.vbxCrystal.Formulas(5) = "group1_comtax=Sum ({HRBENFT.BF_CCOST}, " & grpField$ & ")*({@comptaxperc}/100)"
        
        Me.vbxCrystal.Formulas(6) = "group1_emptotal=Sum ({HRBENFT.BF_ECOST}, " & grpField$ & ")+{@group1_emptax}"
        Me.vbxCrystal.Formulas(7) = "group1_comptotal=Sum ({HRBENFT.BF_CCOST}, " & grpField$ & ")+{@group1_comtax}"
        
        If InStr(1, Me.vbxCrystal.ReportFileName, "rzbenftM.rpt") > 0 Then
            Me.vbxCrystal.Formulas(3) = "group1total=Sum ({HRBENFT.BF_MTHCCOST}, " & grpField$ & ")+Sum ({HRBENFT.BF_MTHECOST}, " & grpField$ & ")"
            
            Me.vbxCrystal.Formulas(4) = "group1_emptax=Sum ({HRBENFT.BF_MTHECOST}, " & grpField$ & ")*({@emptaxperc}/100)"
            Me.vbxCrystal.Formulas(5) = "group1_comtax=Sum ({HRBENFT.BF_MTHCCOST}, " & grpField$ & ")*({@comptaxperc}/100)"
            
            Me.vbxCrystal.Formulas(6) = "group1_emptotal=Sum ({HRBENFT.BF_MTHECOST}, " & grpField$ & ")+{@group1_emptax}"
            Me.vbxCrystal.Formulas(7) = "group1_comptotal=Sum ({HRBENFT.BF_MTHCCOST}, " & grpField$ & ")+{@group1_comtax}"
        End If
    End If
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HRTABL.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        If grpField$ = "{@EFullName}" Then grpField$ = "{HREMP.ED_EMPNBR}"
        
        grpCond$ = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
        
        If optGrouping(1) Then
            Me.vbxCrystal.Formulas(1) = "group2total=Sum ({HRBENFT.BF_CCOST}, " & grpField$ & ")+Sum ({HRBENFT.BF_ECOST}, " & grpField$ & ")"
            
            'Hemu
            If InStr(1, Me.vbxCrystal.ReportFileName, "rzbenftM.rpt") > 0 Then
                'changed by Bryan to calc monthly cost on monthly report. Dec 18/05 Ticket#2375
                Me.vbxCrystal.Formulas(2) = "grouptot=Sum ({HRBENFT.BF_MTHCCOST}, " & grpField$ & ")+Sum ({HRBENFT.BF_MTHECOST}, " & grpField$ & ")"
            End If
            'Hemu
        End If
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HRTABL.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP4;" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(3) = grpCond$
        Else
            'Hemu
            If InStr(1, Me.vbxCrystal.ReportFileName, "rzbnftdb.rpt") = 0 Then
            'Hemu
                grpCond$ = "GROUP4;{HRBENFT.BF_EDATE}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(3) = grpCond$
            End If
        End If
    End If
    'Here sorting starts when first grouping is "none"
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HRTABL.TB_DESC}"
        Case 2: grpField$ = "{HRTABL1.TB_DESC}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        dscGroup$ = comGroup(1).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal.Formulas(0) = dscGroup$
        
        'City of Niagara Falls
        If flgNiagara And glbCompSerial = "S/N - 2276W" Then
            Me.vbxCrystal1.GroupCondition(0) = grpCond$
            Me.vbxCrystal1.Formulas(0) = dscGroup$
        End If
        
        'added by Bryan 16/Dec/05 Ticket#10009
        If optGrouping(1) Then
            Me.vbxCrystal.Formulas(1) = "grpTotal=" & grpField$
        
            'City of Niagara Falls
            If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                Me.vbxCrystal1.Formulas(1) = "grpTotal=" & grpField$
            End If
        End If
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HRTABL.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
         
            'City of Niagara Falls
            If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                Me.vbxCrystal1.GroupCondition(1) = grpCond$
            End If
        End If
    Else
        'Here sorting starts when second grouping is also "none"
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HRTABL.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            dscGroup$ = comGroup(2).Text
            dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
            Me.vbxCrystal.Formulas(0) = dscGroup$
            
            'City of Niagara Falls
            If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                Me.vbxCrystal1.Formulas(0) = dscGroup$
            End If
            
            'added by Bryan 16/Dec/05 Ticket#10009
            If optGrouping(1) Then
                Me.vbxCrystal.Formulas(1) = "grpTotal=" & grpField$
            
                'City of Niagara Falls
                If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                    Me.vbxCrystal1.Formulas(1) = "grpTotal=" & grpField$
                End If
            End If
          
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$

            'City of Niagara Falls
            If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                Me.vbxCrystal1.GroupCondition(0) = grpCond$
            End If
          
            If Not optGrouping(2) Then 'added by RAUBREY 5/14/97
                ' Modified by Frank 04/23/2001 Begin BF_FDATE -> BF_EDATE
                grpCond$ = "GROUP" & CStr(2) & ";" & "{HRBENFT.BF_EDATE}" & ";ANYCHANGE;A"
                ' Modified by Frank 04/23/2001 End
                Me.vbxCrystal.GroupCondition(1) = grpCond$
            
                'City of Niagara Falls
                If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                    Me.vbxCrystal1.GroupCondition(1) = grpCond$
                End If
            End If
            
        'If third grouping is "none"
        Else
            If Not optGrouping(2) Then
                grpCond$ = "GROUP" & CStr(1) & ";" & "{HRBENFT.BF_EDATE}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
            
                'City of Niagara Falls
                If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                    Me.vbxCrystal1.GroupCondition(0) = grpCond$
                    Me.vbxCrystal.GroupCondition(0) = grpCond$
                Else
                    Me.vbxCrystal.GroupCondition(0) = grpCond$
                End If
            
                'added by Bryan 16/Dec/05 Ticket#10009
                If optGrouping(1) Then
                    'Me.vbxCrystal.Formulas(1) = "grpTotal=" & grpField$
                    
                    'City of Niagara Falls
                    If flgNiagara And glbCompSerial = "S/N - 2276W" Then
                        'Me.vbxCrystal1.Formulas(1) = "grpTotal=" & grpField$
                    End If
                End If
            End If
        End If
    End If
    
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = strSFormat$

    'City of Niagara Falls
    If flgNiagara And glbCompSerial = "S/N - 2276W" Then
        Me.vbxCrystal1.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
        strSFormat$ = "GF1;T;X;X;X;X;X;X"
        Me.vbxCrystal1.SectionFormat(1) = strSFormat$
    End If
End If

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%

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


For X% = 0 To 9
If X% <> 8 And X <> 1 And X <> 2 And X <> 3 And X <> 4 And X <> 5 And X <> 6 Then
    If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X%).SetFocus
        Exit Function
    End If
End If
Next X%


For X% = 0 To 1
 If Len(dlpDateRange(X%).Text) > 0 Then
    If Not IsDate(dlpDateRange(X%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(X%).Text = ""
        dlpDateRange(X%).SetFocus
        Exit Function
    End If
 End If
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date cannot be prior to From Date"
        Me.dlpDateRange(0).SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then Exit Function
If Not clpCode(1).ListChecker Then Exit Function
If Not clpCode(5).ListChecker Then Exit Function
If Not clpCode(6).ListChecker Then Exit Function

If optGrouping(1) Then
    If Len(Trim(txtTaxPerc)) > 0 Then
        If Not IsNumeric(txtTaxPerc) Then
            MsgBox "Tax Percentage must be numeric"
            txtTaxPerc.SetFocus
            Exit Function
        End If
    End If
End If

For X% = 0 To 1
 If Len(dlpDate(X%).Text) > 0 Then
    If Not IsDate(dlpDate(X%).Text) Then
        MsgBox "Not a valid date"
        dlpDate(X%).Text = ""
        dlpDate(X%).SetFocus
        Exit Function
    End If
 End If
Next X%

If IsDate(dlpDate(0)) And IsDate(dlpDate(1)) Then
    If DaysBetween(dlpDate(0), dlpDate(1)) < 0 Then
        MsgBox "End To Date cannot be prior to End From Date"
        Me.dlpDate(0).SetFocus
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

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
flgNiagara = False
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRBENEFITS"
Screen.MousePointer = HOURGLASS
Call comGrpLoad
Call setRptCaption(Me)
If glbLinamar Then
'    clpCode(1).MaxLength = 8
'    clpCode(2).MaxLength = 8
'    clpCode(3).MaxLength = 8
'    clpCode(4).MaxLength = 8
    clpCode(9).MaxLength = 8
End If
If glbCompSerial = "S/N - 2227W" Then clpCode(9).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
Call INI_Controls(Me)

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

If optGrouping(1) Then
    txtTaxPerc.Visible = True
    lblTitle(0).Visible = True
End If

If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14189
    lblTitle(1).Caption = lStr("G/L #") 'Ticket #25466 Franks 05/09/2014
    lblTitle(1).Visible = True
    clpGLNum.Visible = True
End If

If Not glbCompSerial = "S/N - 2362W" Then 'City of Sarnia Ticket #23890 Franks 06/21/2013
    optGrouping(3).Enabled = False
End If

Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
scrFrame.Height = 9615
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 10440 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 6500 Then
            scrControl.Max = 5600
        Else
            scrControl.Max = 4100
        End If
        scrControl.Left = Me.Width - scrControl.Width - 250
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 250
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 9600 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 8000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 830
        scrHScroll.Width = Me.Width - 250
    End If
    scrFrame.Refresh
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optAnnMth_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
    If optGrouping(2) Then
        chkDOB.Visible = False
        txtTaxPerc.Visible = False
        lblTitle(0).Visible = False
    ElseIf optGrouping(3) Then 'Ticket #23890 Franks 06/21/2013 - same as (2)
        chkDOB.Visible = False
        txtTaxPerc.Visible = False
        lblTitle(0).Visible = False
    ElseIf optGrouping(1) Then
        chkDOB.Visible = True
        txtTaxPerc.Visible = True
        lblTitle(0).Visible = True
    Else
        chkDOB.Visible = True
        txtTaxPerc.Visible = False
        lblTitle(0).Visible = False
    End If
    If glbCompSerial = "S/N - 2388W" Then 'DNSSAB
        Call comGrpLoad 'Ticket #14189
    End If
    'Ticket #19432 Frank 11/12/2010
    If optGrouping(2) Then
        chkEndedBenefit.Value = 0
        chkEndedBenefit.Enabled = False
    Else
        chkEndedBenefit.Enabled = True
    End If
    'Ticket #19432 - end
End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtCoverageType_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCoverageType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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

Private Sub CitySarniaBillingWrk() 'City of Sarnia Ticket #23890 Franks 07/08/2013
Dim WSQLQ, strSFormat$
Dim DATE1, DATE2
Dim rsEMP As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsEPos As New ADODB.Recordset
Dim rsESal As New ADODB.Recordset
Dim rsEPerf As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ
Dim I As Long
Dim totNum As Long
Dim xSkip As Boolean
Dim xTemp
On Error GoTo Mod_Err
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " Please Wait"
        
    'Delete Temp tables
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    WSQLQ = "WHERE " & glbSeleDeptUn & " "
    If Len(clpDiv.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "') "
    If Len(clpDept.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "') "
    If Len(clpCode(0).Text) > 0 Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "' "
    If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    If Len(clpCode(6).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
    If Len(clpPT.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "') "
    If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    If Len(clpCode(9).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(9).Text, ",", "','") & "') "
    If Len(clpCode(10).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(10).Text, ",", "','") & "') "
    If Len(clpCode(7).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(7).Text, ",", "','") & "') "
    If Len(txtShift.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_SHIFT='" & txtShift.Text & "' "
    'If Len(comCountryOfEmp.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "' "
    SQLQ = "SELECT * FROM HREMP "
    SQLQ = SQLQ & WSQLQ
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    If rsEMP.State <> 0 Then rsEMP.Close
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then totNum = rsEMP.RecordCount
    I = 0
    Do While Not rsEMP.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        xSkip = False
        'do not check Benefit here,
        If Not xSkip Then
            'add to wrk table
            rsWRK.AddNew
            rsWRK("TT_EMPNBR") = rsEMP("ED_EMPNBR")
            rsWRK("TT_NAMEFLD") = Left((Trim(rsEMP("ED_SURNAME")) & " " & Trim(rsEMP("ED_FNAME"))), 40)
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " "
    Exit Sub
Mod_Err:
    MsgBox Err.Description
    'Resume Next
End Sub

Private Sub CitySarniaBillingWrk_TerminatedEmployees() 'City of Sarnia Ticket #24106 Hemu 8/13/2013
Dim WSQLQ, strSFormat$
Dim DATE1, DATE2
Dim rsEMP As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsEPos As New ADODB.Recordset
Dim rsESal As New ADODB.Recordset
Dim rsEPerf As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ
Dim I As Long
Dim totNum As Long
Dim xSkip As Boolean
Dim xTemp
On Error GoTo Mod_Err
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " Please Wait"
        
    'Delete Temp tables
    'SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    'gdbAdoIhr001.Execute SQLQ
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    WSQLQ = "WHERE ED_EMPNBR IN (SELECT EMPLOYEE_NUMBER FROM Term_HRTRMEMP WHERE TERM_REASON IN ('RETA','RETL')) "
    'Ticket #24106 - change in the logic to retrieve Terminated Employees
    'WSQLQ = WSQLQ & " AND ED_UNION >= " & Date_SQL(Format(Now, "mm/dd/yyyy")) & " AND " & glbSeleDeptUn & " "
    WSQLQ = WSQLQ & " AND DATEDIFF(dd, ED_DOB, GETDATE())/365 < 65 AND (ED_UNION = '' OR ED_UNION IS NULL) AND " & glbSeleDeptUn & " "
    
    If Len(clpDiv.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "') "
    If Len(clpDept.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "') "
    If Len(clpCode(0).Text) > 0 Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "' "
    If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    If Len(clpCode(6).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
    If Len(clpPT.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "') "
    
    'This is getting employees from Term table and the selection criteria is dependent on Active employees.
    If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    
    If Len(clpCode(9).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(9).Text, ",", "','") & "') "
    If Len(clpCode(10).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(10).Text, ",", "','") & "') "
    If Len(clpCode(7).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(7).Text, ",", "','") & "') "
    If Len(txtShift.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_SHIFT='" & txtShift.Text & "' "
    'If Len(comCountryOfEmp.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "' "
    
    SQLQ = "SELECT * FROM Term_HREMP "
    SQLQ = SQLQ & WSQLQ
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    
    If rsEMP.State <> 0 Then rsEMP.Close
    rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    If Not rsEMP.EOF Then totNum = rsEMP.RecordCount
    I = 0
    Do While Not rsEMP.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        xSkip = False
        'do not check Benefit here,
        If Not xSkip Then
            'add to wrk table
            rsWRK.AddNew
            rsWRK("TT_EMPNBR") = rsEMP("ED_EMPNBR")
            rsWRK("TT_NAMEFLD") = Left((Trim(rsEMP("ED_SURNAME")) & " " & Trim(rsEMP("ED_FNAME"))), 40)
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK("TT_FLAG") = "Y"
            rsWRK("TT_REPTAU") = rsEMP("TERM_SEQ")  'Using this field to store the TERM_SEQ #. This is just a temp table and this field is not being used.
            rsWRK.Update
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " "
    Exit Sub
Mod_Err:
    MsgBox Err.Description
    'Resume Next
End Sub

'Private Function CitySarniaGenerate_Excel_File() 'City of Sarnia Ticket #23890 Franks 07/08/2013
'    Dim rsEMP As New ADODB.Recordset
'    Dim rsWRK As New ADODB.Recordset
'    Dim rsESal As New ADODB.Recordset
'    Dim rsEBen As New ADODB.Recordset
'    Dim SQLQ As String
'    Dim xEmpNo As String
'    Dim xlsFileMat As String
'    Dim xlsFileMain As String
'    Dim xlsFileTmp As String
'    Dim xTrainMatrixPath
'    Dim xSaveFileName As String
'    Dim xRow, xRecNum As Integer
'    Dim strReptAuth1 As String
'    Dim I As Long
'    Dim totNum As Long
'    Dim k As Long
'    Dim EmpName, EmpDocName
'    Dim xFName, xTemp
'    Dim locPath, xLocPath
'    Dim xSaveToPath
'    Dim strRName$
'    Dim xBenList
'    Dim exApp As Object, exBook As Object, exSheet As Object
'    Dim xLF1, xLF2
'    Dim xDEP1, xDEP2
'    Dim xWI1, xWI2
'    Dim xLTD1, xLTD2
'    Dim xAD1, xAD2
'
'    '''On Error GoTo Err_Word_Forms
'
'    Screen.MousePointer = HOURGLASS
'
'    xBenList = "(LEFT(BF_BCODE,2)= 'LF' OR LEFT(BF_BCODE,3)= 'DEP' OR LEFT(BF_BCODE,2)= 'WI' OR LEFT(BF_BCODE,3)= 'LTD' OR LEFT(BF_BCODE,2)= 'AD') "
'
'    MDIMain.panHelp(0).FloodType = 1
'    MDIMain.panHelp(0).FloodPercent = 0
'    MDIMain.panHelp(0).Caption = "Please wait..."
'
'    'SQLQ = "DELETE FROM PAYROLL_TRANSFER_RPT WHERE [TABLE] = 'SarniaBilling'"
'    'gdbAdoIhr001.Execute SQLQ
'    'Call Pause(1)
'
'    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_NAMEFLD " ' TT_ID"
'    If rsWRK.State <> 0 Then rsWRK.Close
'    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If rsWRK.EOF Then
'        MsgBox "No record found."
'        Exit Function
'    End If
'    I = 0
'    If Not rsWRK.EOF Then totNum = rsWRK.RecordCount
'
'    'Begin ------------------------------------------------------------------
'    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "City of Sarnia Benefit Billing Report_Tmp.xlsx"
'    If gsTRAININGMATRIX Then
'        xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
'    End If
'    If Len(xTrainMatrixPath) = 0 Then
'        xTrainMatrixPath = glbIHRREPORTS
'    End If
'    xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "City of Sarnia Benefit Billing Report" & Trim(glbUserID) & ".xlsx"
'
'    MDIMain.panHelp(0).FloodType = 1
'    MDIMain.panHelp(0).FloodPercent = 0
'    MDIMain.panHelp(0).Caption = "Please wait..."
'
'    If Dir(xlsFileTmp) = "" Then
'        MsgBox "There is no " & xlsFileTmp
'        Exit Function
'    End If
'    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
'
'    FileCopy xlsFileTmp, xlsFileMat
'
'    'Create new WorkBook of Excel
'    Set exApp = New Excel.Application
'    Set exBook = exApp.Workbooks.Open(xlsFileMat)
'    Set exSheet = exBook.Worksheets(1)
'    xTemp = ""
'    If IsDate(dlpDateRange(0).Text) Then
'        xTemp = xTemp & Format(dlpDateRange(0).Text, "dd-mmm-yyyy")
'    End If
'    If IsDate(dlpDateRange(1).Text) Then
'        xTemp = xTemp & " to " & Format(dlpDateRange(1).Text, "dd-mmm-yyyy")
'    End If
'    exSheet.Cells(3, 2) = xTemp
'    k = 5
'    Do While Not rsWRK.EOF
'        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
'        I = I + 1
'        DoEvents
'
'        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & rsWRK("TT_EMPNBR") & " "
'        If rsEMP.State <> 0 Then rsEMP.Close
'        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
'        If Not rsEMP.EOF Then
'            EmpName = rsEMP("ED_SURNAME") & ", " & rsEMP("ED_FNAME")
'            EmpDocName = rsEMP("ED_FNAME") & " " & rsEMP("ED_SURNAME")
'
'            'check if this employee has these benefits, if not then skip
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND " & xBenList & " "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            If rsEBen.EOF Then
'                GoTo next_rec
'            End If
'            rsEBen.Close
'
'
'            k = k + 1
'            exSheet.Cells(k, 1) = EmpName ' EmpDocName
'            exSheet.Cells(k, 2) = rsEMP("ED_EMPNBR")
'            exSheet.Cells(k, 4) = getDeptDesc(rsEMP("ED_DEPTNO"))
'            If IsNull(rsEMP("ED_DOB")) Then xTemp = "" Else xTemp = CVDate(rsEMP("ED_DOB")) ' Format(rsEMP("ED_DOB"), "MMM D, YYYY")
'            exSheet.Cells(k, 5) = xTemp
'            If IsNull(rsEMP("ED_DOH")) Then xTemp = "" Else xTemp = CVDate(rsEMP("ED_DOH")) ' Format(rsEMP("ED_DOH"), "MMM D, YYYY")
'            exSheet.Cells(k, 6) = xTemp
'
'            'Get Salary
'            xTemp = ""
'            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            If rsESal.State <> 0 Then rsESal.Close
'            rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            If Not rsESal.EOF Then
'                xTemp = rsESal("SH_SALARY")
'                exSheet.Cells(k, 7) = xTemp
'                If Not IsNull(rsESal("SH_WHRS")) Then exSheet.Cells(k, 8) = rsESal("SH_WHRS")
'            End If
'            rsESal.Close
'
'            'Benefit Group - begin
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND " & xBenList & " "
'            SQLQ = SQLQ & "AND NOT (BF_GROUP IS NULL) "
'            SQLQ = SQLQ & "ORDER BY BF_GROUP, BF_EDATE DESC"
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            If Not rsEBen.EOF Then
'                xTemp = rsEBen("BF_GROUP")
'                exSheet.Cells(k, 3) = GetTABLDesc("BGMF", xTemp)
'            End If
'            rsEBen.Close
'            'Benefit Group - end
'
'            'LF Benefit - begin
'            'Coverage Amount for Benefit Codes beginning with "LF".
'            xLF1 = 0: xLF2 = 0
'            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'LF' "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            Do While Not rsEBen.EOF
'                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
'                    If IsDate(dlpDateRange(0).Text) Then
'                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
'                            GoTo next_LF 'skip
'                        End If
'                    End If
'                End If
'                If Not IsNull(rsEBen("BF_AMT")) Then
'                    'exSheet.Cells(K, 9) = rsEBen("TOT_AMT")
'                    xLF1 = xLF1 + rsEBen("BF_AMT")
'                End If
'                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
'                    'exSheet.Cells(K, 10) = rsEBen("TOT_MTHCCOST")
'                    xLF2 = xLF2 + rsEBen("BF_MTHCCOST")
'                End If
'next_LF:
'                rsEBen.MoveNext
'            Loop
'            rsEBen.Close
'            If xLF1 > 0 Then exSheet.Cells(k, 9) = xLF1
'            If xLF2 > 0 Then exSheet.Cells(k, 10) = xLF2
'            'LF Benefit - end
'
'            'DEP Benefit - begin
'            'Coverage Amount for Benefit Codes beginning with "DEP".
'            xDEP1 = 0: xDEP2 = 0
'            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND LEFT(BF_BCODE,3)= 'DEP' "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            'If Not rsEBen.EOF Then
'            Do While Not rsEBen.EOF
'                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
'                    If IsDate(dlpDateRange(0).Text) Then
'                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
'                            GoTo next_DEP 'skip
'                        End If
'                    End If
'                End If
'                If Not IsNull(rsEBen("BF_AMT")) Then
'                    'exSheet.Cells(K, 11) = rsEBen("TOT_AMT")
'                    xDEP1 = xDEP1 + rsEBen("BF_AMT")
'                End If
'                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
'                    'exSheet.Cells(K, 12) = rsEBen("TOT_MTHCCOST")
'                    xDEP2 = xDEP2 + rsEBen("BF_MTHCCOST")
'                End If
'next_DEP:
'                rsEBen.MoveNext
'            Loop
'            'End If
'            rsEBen.Close
'            If xDEP1 > 0 Then exSheet.Cells(k, 11) = xDEP1
'            If xDEP2 > 0 Then exSheet.Cells(k, 12) = xDEP2
'            'DEP Benefit - end
'
'            'WI Benefit - begin
'            'Coverage Amount for Benefit Codes beginning with "WI".
'            xWI1 = 0: xWI2 = 0
'            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'WI' "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            'If Not rsEBen.EOF Then
'            Do While Not rsEBen.EOF
'                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
'                    If IsDate(dlpDateRange(0).Text) Then
'                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
'                            GoTo next_WI 'skip
'                        End If
'                    End If
'                End If
'                If Not IsNull(rsEBen("BF_AMT")) Then
'                    'exSheet.Cells(K, 13) = rsEBen("TOT_AMT")
'                    xWI1 = xWI1 + rsEBen("BF_AMT")
'                End If
'                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
'                    'exSheet.Cells(K, 14) = rsEBen("TOT_MTHCCOST")
'                    xWI2 = xWI2 + rsEBen("BF_MTHCCOST")
'                End If
'next_WI:
'                rsEBen.MoveNext
'            Loop
'            'End If
'            rsEBen.Close
'            If xWI1 > 0 Then exSheet.Cells(k, 13) = xWI1
'            If xWI2 > 0 Then exSheet.Cells(k, 14) = xWI2
'            'WI Benefit - end
'
'            'LTD Benefit - begin
'            'Coverage Amount for Benefit Codes beginning with "LTD".
'            xLTD1 = 0: xLTD2 = 0
'            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND LEFT(BF_BCODE,3)= 'LTD' "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            'If Not rsEBen.EOF Then
'            Do While Not rsEBen.EOF
'                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
'                    If IsDate(dlpDateRange(0).Text) Then
'                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
'                            GoTo next_LTD 'skip
'                        End If
'                    End If
'                End If
'                If Not IsNull(rsEBen("BF_AMT")) Then
'                    'exSheet.Cells(K, 15) = rsEBen("TOT_AMT")
'                    xLTD1 = xLTD1 + rsEBen("BF_AMT")
'                End If
'                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
'                    'exSheet.Cells(K, 16) = rsEBen("TOT_MTHCCOST")
'                    xLTD2 = xLTD2 + rsEBen("BF_MTHCCOST")
'                End If
'next_LTD:
'                rsEBen.MoveNext
'            Loop
'            'End If
'            rsEBen.Close
'            If xLTD1 > 0 Then exSheet.Cells(k, 15) = xLTD1
'            If xLTD2 > 0 Then exSheet.Cells(k, 16) = xLTD2
'            'LTD Benefit - end
'
'            'AD Benefit - begin
'            'Coverage Amount for Benefit Codes beginning with "AD".
'            xAD1 = 0: xAD2 = 0
'            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
'            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'AD' "
'            If IsDate(dlpDateRange(0).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
'            End If
'            If IsDate(dlpDateRange(1).Text) Then
'                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
'            End If
'            If IsDate(dlpDate(0).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
'            End If
'            If IsDate(dlpDate(1).Text) Then
'                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
'            End If
'            If Len(txtCoverageType.Text) > 0 Then
'                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
'            End If
'            If rsEBen.State <> 0 Then rsEBen.Close
'            rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            'If Not rsEBen.EOF Then
'             Do While Not rsEBen.EOF
'                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
'                    If IsDate(dlpDateRange(0).Text) Then
'                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
'                            GoTo next_AD 'skip
'                        End If
'                    End If
'                End If
'                If Not IsNull(rsEBen("BF_AMT")) Then
'                    'exSheet.Cells(K, 17) = rsEBen("TOT_AMT")
'                    xAD1 = xAD1 + rsEBen("BF_AMT")
'                End If
'                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
'                    'exSheet.Cells(K, 18) = rsEBen("TOT_MTHCCOST")
'                    xAD2 = xAD2 + rsEBen("BF_MTHCCOST")
'                End If
'next_AD:
'                rsEBen.MoveNext
'            Loop
'            'End If
'            rsEBen.Close
'            If xAD1 > 0 Then exSheet.Cells(k, 17) = xAD1
'            If xAD2 > 0 Then exSheet.Cells(k, 18) = xAD2
'            'AD Benefit - end
'
'            'Total for all benefits (including tax) - column S
'            xTemp = xLF2 + xDEP2 + xWI2 + xLTD2 + xAD2
'            'Ticket #24106 - Add the Tax
'            'exSheet.Cells(k, 19) = xTemp
'            exSheet.Cells(k, 19) = xTemp * 1.08
'
'            'Tax
'            exSheet.Cells(k, 20) = xTemp * 0.08
'        End If
'next_rec:
'        rsWRK.MoveNext
'    Loop
'    rsWRK.Close
'
'    exBook.Save
'    Set exSheet = Nothing
'    Set exBook = Nothing
'    exApp.Quit
'    Set exApp = Nothing
'
'    Call Pause(1)
'
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(0).Caption = ""
'    MDIMain.panHelp(1).Caption = ""
'    MDIMain.panHelp(2).Caption = ""
'
'    MDIMain.Timer1.Enabled = False
'    Screen.MousePointer = DEFAULT
'
'    If Not LanchXlsW98(xlsFileMat) Then
'        Shell "cmd /c " & GetShortName(xlsFileMat)
'    End If
'
'Exit Function
'
'Err_Word_Forms:
'
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(0).Caption = ""
'    MDIMain.panHelp(1).Caption = ""
'    MDIMain.panHelp(2).Caption = ""
'
'    Screen.MousePointer = DEFAULT
'
'    If Err = 75 Then
'        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
'        Exit Function
'    End If
'    If Not exBook Is Nothing Then
'        exBook.Save
'        Set exSheet = Nothing
'        Set exBook = Nothing
'        exApp.Quit
'    End If
'    If Err = 70 Then
'        Set exApp = Nothing
'        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
'        Exit Function
'    End If
'
'    glbFrmCaption$ = Me.Caption
'    glbErrNum& = Err
'    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CitySarniaGenerate_Excel_File", "", "Select")
'
'End Function

Private Function CitySarniaGenerate_Excel_File() 'City of Sarnia Ticket #23890 Franks 07/08/2013
    Dim rsEMP As New ADODB.Recordset
    Dim rsWRK As New ADODB.Recordset
    Dim rsESal As New ADODB.Recordset
    Dim rsEBen As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEmpNo As String
    Dim xlsFileMat As String
    Dim xlsFileMain As String
    Dim xlsFileTmp As String
    Dim xTrainMatrixPath
    Dim xSaveFileName As String
    Dim xRow, xRecNum As Integer
    Dim strReptAuth1 As String
    Dim I As Long
    Dim totNum As Long
    Dim K As Long
    Dim EmpName, EmpDocName
    Dim xFName, xTemp
    Dim locPath, xLocPath
    Dim xSaveToPath
    Dim strRName$
    Dim xBenList
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim xLF1, xLF2
    Dim xDEP1, xDEP2
    Dim xWI1, xWI2
    Dim xLTD1, xLTD2
    Dim xAD1, xAD2
    
    '''On Error GoTo Err_Word_Forms
    
    Screen.MousePointer = HOURGLASS
    
    xBenList = "(LEFT(BF_BCODE,2)= 'LF' OR LEFT(BF_BCODE,3)= 'DEP' OR LEFT(BF_BCODE,2)= 'WI' OR LEFT(BF_BCODE,3)= 'LTD' OR LEFT(BF_BCODE,2)= 'AD') "
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."
    
    'SQLQ = "DELETE FROM PAYROLL_TRANSFER_RPT WHERE [TABLE] = 'SarniaBilling'"
    'gdbAdoIhr001.Execute SQLQ
    'Call Pause(1)
    
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_NAMEFLD " ' TT_ID"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsWRK.EOF Then
        MsgBox "No record found."
        Exit Function
    End If
    I = 0
    If Not rsWRK.EOF Then totNum = rsWRK.RecordCount
    
    'Begin ------------------------------------------------------------------
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "City of Sarnia Benefit Billing Report_Tmp.xlsx"
    If gsTRAININGMATRIX Then
        xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xTrainMatrixPath) = 0 Then
        xTrainMatrixPath = glbIHRREPORTS
    End If
    xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "City of Sarnia Benefit Billing Report" & Trim(glbUserID) & ".xlsx"

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Function
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
    'Create new WorkBook of Excel
    Set exApp = New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    xTemp = ""
    If IsDate(dlpDateRange(0).Text) Then
        xTemp = xTemp & Format(dlpDateRange(0).Text, "dd-mmm-yyyy")
    End If
    If IsDate(dlpDateRange(1).Text) Then
        xTemp = xTemp & " to " & Format(dlpDateRange(1).Text, "dd-mmm-yyyy")
    End If
    exSheet.Cells(3, 2) = xTemp
    K = 5
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        
        'Ticket #24106 - If Terminated Employees
        If rsWRK("TT_FLAG") = "Y" Then
            SQLQ = "SELECT * FROM TERM_HREMP WHERE ED_EMPNBR = " & rsWRK("TT_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            If rsEMP.State <> 0 Then rsEMP.Close
            rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        Else
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & rsWRK("TT_EMPNBR") & " "
            If rsEMP.State <> 0 Then rsEMP.Close
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
        End If
        If Not rsEMP.EOF Then
            EmpName = rsEMP("ED_SURNAME") & ", " & rsEMP("ED_FNAME")
            EmpDocName = rsEMP("ED_FNAME") & " " & rsEMP("ED_SURNAME")
            
            'check if this employee has these benefits, if not then skip
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND " & xBenList & " "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            If rsEBen.EOF Then
                GoTo next_rec
            End If
            rsEBen.Close
            
            
            K = K + 1
            exSheet.Cells(K, 1) = EmpName ' EmpDocName
            exSheet.Cells(K, 2) = rsEMP("ED_EMPNBR")
            exSheet.Cells(K, 4) = getDeptDesc(rsEMP("ED_DEPTNO"))
            If IsNull(rsEMP("ED_DOB")) Then xTemp = "" Else xTemp = CVDate(rsEMP("ED_DOB")) ' Format(rsEMP("ED_DOB"), "MMM D, YYYY")
            exSheet.Cells(K, 5) = xTemp
            If IsNull(rsEMP("ED_DOH")) Then xTemp = "" Else xTemp = CVDate(rsEMP("ED_DOH")) ' Format(rsEMP("ED_DOH"), "MMM D, YYYY")
            exSheet.Cells(K, 6) = xTemp
            
            'Get Salary
            xTemp = ""
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
                If rsESal.State <> 0 Then rsESal.Close
                rsESal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
                If rsESal.State <> 0 Then rsESal.Close
                rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            If Not rsESal.EOF Then
                xTemp = rsESal("SH_SALARY")
                exSheet.Cells(K, 7) = xTemp
                If Not IsNull(rsESal("SH_WHRS")) Then exSheet.Cells(K, 8) = rsESal("SH_WHRS")
            End If
            rsESal.Close
            
            'Benefit Group - begin
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
                SQLQ = SQLQ & "AND " & xBenList & " "
                SQLQ = SQLQ & "AND NOT (BF_GROUP IS NULL) "
                SQLQ = SQLQ & "ORDER BY BF_GROUP, BF_EDATE DESC"
                If rsEBen.State <> 0 Then rsEBen.Close
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND " & xBenList & " "
                SQLQ = SQLQ & "AND NOT (BF_GROUP IS NULL) "
                SQLQ = SQLQ & "ORDER BY BF_GROUP, BF_EDATE DESC"
                If rsEBen.State <> 0 Then rsEBen.Close
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            If Not rsEBen.EOF Then
                xTemp = rsEBen("BF_GROUP")
                exSheet.Cells(K, 3) = GetTABLDesc("BGMF", xTemp)
            End If
            rsEBen.Close
            'Benefit Group - end
            
            'LF Benefit - begin
            'Coverage Amount for Benefit Codes beginning with "LF".
            xLF1 = 0: xLF2 = 0
            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'LF' "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            Do While Not rsEBen.EOF
                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
                            GoTo next_LF 'skip
                        End If
                    End If
                End If
                If Not IsNull(rsEBen("BF_AMT")) Then
                    'exSheet.Cells(K, 9) = rsEBen("TOT_AMT")
                    xLF1 = xLF1 + rsEBen("BF_AMT")
                End If
                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
                    'exSheet.Cells(K, 10) = rsEBen("TOT_MTHCCOST")
                    xLF2 = xLF2 + rsEBen("BF_MTHCCOST")
                End If
next_LF:
                rsEBen.MoveNext
            Loop
            rsEBen.Close
            If xLF1 > 0 Then exSheet.Cells(K, 9) = xLF1
            If xLF2 > 0 Then exSheet.Cells(K, 10) = xLF2
            'LF Benefit - end
            
            'DEP Benefit - begin
            'Coverage Amount for Benefit Codes beginning with "DEP".
            xDEP1 = 0: xDEP2 = 0
            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND LEFT(BF_BCODE,3)= 'DEP' "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            'If Not rsEBen.EOF Then
            Do While Not rsEBen.EOF
                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
                            GoTo next_DEP 'skip
                        End If
                    End If
                End If
                If Not IsNull(rsEBen("BF_AMT")) Then
                    'exSheet.Cells(K, 11) = rsEBen("TOT_AMT")
                    xDEP1 = xDEP1 + rsEBen("BF_AMT")
                End If
                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
                    'exSheet.Cells(K, 12) = rsEBen("TOT_MTHCCOST")
                    xDEP2 = xDEP2 + rsEBen("BF_MTHCCOST")
                End If
next_DEP:
                rsEBen.MoveNext
            Loop
            'End If
            rsEBen.Close
            If xDEP1 > 0 Then exSheet.Cells(K, 11) = xDEP1
            If xDEP2 > 0 Then exSheet.Cells(K, 12) = xDEP2
            'DEP Benefit - end
            
            'WI Benefit - begin
            'Coverage Amount for Benefit Codes beginning with "WI".
            xWI1 = 0: xWI2 = 0
            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'WI' "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            'If Not rsEBen.EOF Then
            Do While Not rsEBen.EOF
                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
                            GoTo next_WI 'skip
                        End If
                    End If
                End If
                If Not IsNull(rsEBen("BF_AMT")) Then
                    'exSheet.Cells(K, 13) = rsEBen("TOT_AMT")
                    xWI1 = xWI1 + rsEBen("BF_AMT")
                End If
                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
                    'exSheet.Cells(K, 14) = rsEBen("TOT_MTHCCOST")
                    xWI2 = xWI2 + rsEBen("BF_MTHCCOST")
                End If
next_WI:
                rsEBen.MoveNext
            Loop
            'End If
            rsEBen.Close
            If xWI1 > 0 Then exSheet.Cells(K, 13) = xWI1
            If xWI2 > 0 Then exSheet.Cells(K, 14) = xWI2
            'WI Benefit - end
            
            'LTD Benefit - begin
            'Coverage Amount for Benefit Codes beginning with "LTD".
            xLTD1 = 0: xLTD2 = 0
            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND LEFT(BF_BCODE,3)= 'LTD' "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            'If Not rsEBen.EOF Then
            Do While Not rsEBen.EOF
                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
                            GoTo next_LTD 'skip
                        End If
                    End If
                End If
                If Not IsNull(rsEBen("BF_AMT")) Then
                    'exSheet.Cells(K, 15) = rsEBen("TOT_AMT")
                    xLTD1 = xLTD1 + rsEBen("BF_AMT")
                End If
                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
                    'exSheet.Cells(K, 16) = rsEBen("TOT_MTHCCOST")
                    xLTD2 = xLTD2 + rsEBen("BF_MTHCCOST")
                End If
next_LTD:
                rsEBen.MoveNext
            Loop
            'End If
            rsEBen.Close
            If xLTD1 > 0 Then exSheet.Cells(K, 15) = xLTD1
            If xLTD2 > 0 Then exSheet.Cells(K, 16) = xLTD2
            'LTD Benefit - end
            
            'AD Benefit - begin
            'Coverage Amount for Benefit Codes beginning with "AD".
            xAD1 = 0: xAD2 = 0
            'SQLQ = "SELECT SUM(BF_AMT) AS TOT_AMT, SUM(BF_MTHCCOST) AS TOT_MTHCCOST FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                SQLQ = "SELECT * FROM TERM_HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND TERM_SEQ = " & rsWRK("TT_REPTAU") & " "
            Else
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsEMP("ED_EMPNBR") & " "
            End If
            SQLQ = SQLQ & "AND LEFT(BF_BCODE,2)= 'AD' "
            If IsDate(dlpDateRange(0).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
            End If
            If IsDate(dlpDateRange(1).Text) Then
                SQLQ = SQLQ & "AND BF_EDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
            End If
            If IsDate(dlpDate(0).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE >= " & Date_SQL(dlpDate(0).Text) & " "
            End If
            If IsDate(dlpDate(1).Text) Then
                SQLQ = SQLQ & "AND BF_CEASEDATE <= " & Date_SQL(dlpDate(1).Text) & " "
            End If
            If Len(txtCoverageType.Text) > 0 Then
                SQLQ = SQLQ & "AND BF_COVER = '" & txtCoverageType.Text & "' "
            End If
            
            'Ticket #24106 - Exclude Ended Benefit logic
            If chkEndedBenefit.Value = 1 Then
                SQLQ = SQLQ & "AND ((BF_CEASEDATE IS NULL) OR BF_CEASEDATE > " & Date_SQL(Format(Now, "mm/dd/yyyy")) & ") "
            End If
            
            If rsEBen.State <> 0 Then rsEBen.Close
            'Ticket #24106 - If Terminated Employees
            If rsWRK("TT_FLAG") = "Y" Then
                rsEBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            Else
                rsEBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            End If
            'If Not rsEBen.EOF Then
             Do While Not rsEBen.EOF
                If Not IsNull(rsEBen("BF_CEASEDATE")) Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If CVDate(rsEBen("BF_CEASEDATE")) < CVDate(dlpDateRange(0).Text) Then
                            GoTo next_AD 'skip
                        End If
                    End If
                End If
                If Not IsNull(rsEBen("BF_AMT")) Then
                    'exSheet.Cells(K, 17) = rsEBen("TOT_AMT")
                    xAD1 = xAD1 + rsEBen("BF_AMT")
                End If
                If Not IsNull(rsEBen("BF_MTHCCOST")) Then
                    'exSheet.Cells(K, 18) = rsEBen("TOT_MTHCCOST")
                    xAD2 = xAD2 + rsEBen("BF_MTHCCOST")
                End If
next_AD:
                rsEBen.MoveNext
            Loop
            'End If
            rsEBen.Close
            If xAD1 > 0 Then exSheet.Cells(K, 17) = xAD1
            If xAD2 > 0 Then exSheet.Cells(K, 18) = xAD2
            'AD Benefit - end
            
            'Total for all benefits (including tax) - column S
            xTemp = xLF2 + xDEP2 + xWI2 + xLTD2 + xAD2
            'Ticket #24106 - Add the Tax
            'exSheet.Cells(K, 19) = xTemp
            exSheet.Cells(K, 19) = xTemp * 1.08
            
            'Tax
            exSheet.Cells(K, 20) = xTemp * 0.08
        End If
next_rec:
        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    Call Pause(1)
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT

    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
Exit Function

Err_Word_Forms:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
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

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CitySarniaGenerate_Excel_File", "", "Select")

End Function

Private Function getDeptDesc(xCode) 'City of Sarnia Ticket #23890 Franks 07/08/2013
Dim rsDEPT As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRDEPT WHERE DF_NBR = '" & xCode & "' "
        rsDEPT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDEPT.EOF Then
            xRetVal = rsDEPT("DF_NAME")
        End If
        rsDEPT.Close
    End If
    getDeptDesc = xRetVal
End Function

