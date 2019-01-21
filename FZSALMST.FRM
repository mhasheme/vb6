VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRSALMAST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary Master Report"
   ClientHeight    =   11175
   ClientLeft      =   390
   ClientTop       =   1230
   ClientWidth     =   13245
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11175
   ScaleWidth      =   13245
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   10335
      LargeChange     =   315
      Left            =   12720
      Max             =   100
      SmallChange     =   315
      TabIndex        =   91
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox pcSalary 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   120
      ScaleHeight     =   10695
      ScaleWidth      =   12495
      TabIndex        =   54
      Top             =   120
      Width           =   12495
      Begin VB.CheckBox chkIncIndSal 
         Caption         =   "Include Indirect Salary"
         Height          =   285
         Left            =   3600
         TabIndex        =   50
         Tag             =   "Check to Include Indirect Salary"
         Top             =   8280
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CheckBox chkExclRET 
         Caption         =   "Exclude Employment Status of RET"
         Height          =   285
         Left            =   135
         TabIndex        =   49
         Tag             =   "Check to Exclude Employees with RET Employment Status"
         Top             =   8580
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.CheckBox chkExclCONP 
         Caption         =   "Exclude Employment Status of CONP"
         Height          =   285
         Left            =   135
         TabIndex        =   48
         Tag             =   "Check to Exclude Employees with CONP Employment Status"
         Top             =   8280
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Tag             =   "First Level of grouping records"
         Top             =   9495
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Tag             =   "Second level of grouping records"
         Top             =   9840
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Tag             =   "Final Sort of Records"
         Top             =   10200
         Width           =   2325
      End
      Begin VB.CheckBox chkVDesc 
         Caption         =   "View Description"
         Height          =   375
         Left            =   7230
         TabIndex        =   38
         Top             =   6795
         Width           =   1605
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   55
         Top             =   6420
         Width           =   4425
         Begin Threed.SSOption optGrouping 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Tag             =   "Type of Report"
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "   Salary"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
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
            Left            =   1320
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "Type of Report"
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "   Position"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optGrouping 
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   33
            TabStop         =   0   'False
            Tag             =   "Type of Report"
            Top             =   60
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "   Backup Position"
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
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   18
         Tag             =   "00-Employee Position Shift"
         Top             =   4046
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtBand 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "JB_BAND"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   11400
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox cmbBand 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11400
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "00-Band"
         Top             =   3685
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkActing 
         Caption         =   "Acting Positions"
         Height          =   225
         Left            =   8910
         TabIndex        =   43
         Top             =   7230
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox comCountryOfEmp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8400
         TabIndex        =   13
         Tag             =   "00-Country of Employment"
         Top             =   3339
         Width           =   1440
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         Height          =   285
         Left            =   1770
         TabIndex        =   8
         Tag             =   "00-Enter Position Code"
         Top             =   2662
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
         Index           =   5
         Left            =   5790
         TabIndex        =   37
         Top             =   6840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   4470
         TabIndex        =   36
         Top             =   6840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   3210
         TabIndex        =   35
         Top             =   6840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1785
         TabIndex        =   34
         Top             =   6840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   1770
         TabIndex        =   9
         Tag             =   "00-Position Group  Code"
         Top             =   3008
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin Threed.SSFrame fraCodeDesc 
         Height          =   1455
         Left            =   9360
         TabIndex        =   57
         Top             =   150
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   2566
         _StockProps     =   14
         Caption         =   "Reason for Change Description"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   61
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   60
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   59
            Top             =   870
            Width           =   840
         End
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   58
            Top             =   1140
            Width           =   840
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1770
         TabIndex        =   5
         Tag             =   "00-Enter Status Code"
         Top             =   1624
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
         Left            =   1770
         TabIndex        =   6
         Tag             =   "EDPT-Category"
         Top             =   1970
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
         Left            =   1770
         TabIndex        =   4
         Tag             =   "00-Enter Union Code"
         Top             =   1278
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
         Left            =   1770
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   932
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1770
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   586
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
         Left            =   1770
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
         Index           =   9
         Left            =   1770
         TabIndex        =   14
         Tag             =   "00-Enter Administered By Code"
         Top             =   3700
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   1770
         TabIndex        =   17
         Tag             =   "00-Enter Section Code"
         Top             =   4046
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   1770
         TabIndex        =   12
         Tag             =   "00-Enter Region Code"
         Top             =   3354
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   45
         Tag             =   "40-Date upto and including this date forward"
         Top             =   7545
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1785
         TabIndex        =   44
         Tag             =   "40-Date from and including this date forward"
         Top             =   7545
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1770
         TabIndex        =   7
         Tag             =   "10-Enter Employee Number"
         Top             =   2316
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   0
         Left            =   1785
         TabIndex        =   39
         Tag             =   "00-Reporting Authority 1"
         Top             =   7200
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         ShowDescription =   0   'False
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   40
         Tag             =   "00-Reporting Authority 2"
         Top             =   7200
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         ShowDescription =   0   'False
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Tag             =   "00-Reporting Authority 3"
         Top             =   7200
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   3
         Left            =   3480
         TabIndex        =   23
         Tag             =   "40-Date upto and including this date forward"
         Top             =   5040
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   2
         Left            =   1770
         TabIndex        =   22
         Tag             =   "40-Date from and including this date forward"
         Top             =   5040
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         DataField       =   "JB_GRID"
         Height          =   285
         Left            =   8100
         TabIndex        =   11
         Top             =   3008
         Visible         =   0   'False
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGD"
      End
      Begin INFOHR_Controls.CodeLookup clpGLNum 
         Height          =   285
         Left            =   8100
         TabIndex        =   20
         Tag             =   "00-General Ledger - Code"
         Top             =   4385
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   1770
         TabIndex        =   24
         Tag             =   "31-Province of Residence - Code"
         Top             =   5380
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin INFOHR_Controls.CodeLookup clpProvEmp 
         Height          =   285
         Left            =   1770
         TabIndex        =   25
         Tag             =   "31-Province of Employment - Code"
         Top             =   5730
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   56
         Top             =   6000
         Width           =   7875
         Begin INFOHR_Controls.DateLookup dlpDateAs 
            Height          =   285
            Left            =   4560
            TabIndex        =   30
            Tag             =   "40-As of date"
            Top             =   120
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin Threed.SSOption optCurrHis 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Tag             =   "Type of Report"
            Top             =   120
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "   Current"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optCurrHis 
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   28
            TabStop         =   0   'False
            Tag             =   "Type of Report"
            Top             =   120
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "   History"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optCurrHis 
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "Type of Report"
            Top             =   120
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "   As of Date (< or = to)"
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
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   11
         Left            =   1770
         TabIndex        =   19
         Tag             =   "00-Supervisory Code for cheque sorting "
         Top             =   4385
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSP"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim1 
         Height          =   285
         Left            =   1770
         TabIndex        =   21
         Top             =   4710
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   12
         Left            =   6480
         TabIndex        =   3
         Tag             =   "00-Enter Physical Branch Code"
         Top             =   932
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SUDE"
      End
      Begin INFOHR_Controls.DateLookup dlpSenDateRange 
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   47
         Tag             =   "40-Date upto and including this date forward"
         Top             =   7890
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1225
      End
      Begin INFOHR_Controls.DateLookup dlpSenDateRange 
         Height          =   285
         Index           =   0
         Left            =   1785
         TabIndex        =   46
         Tag             =   "40-Date from and including this date forward"
         Top             =   7890
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1225
      End
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   3
         Left            =   7080
         TabIndex        =   42
         Tag             =   "00-Reporting Authority 4"
         Top             =   7200
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   13
         Left            =   8100
         TabIndex        =   15
         Tag             =   "00-Band - Code"
         Top             =   3705
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "WFBD"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpJobMaster 
         Height          =   285
         Left            =   8100
         TabIndex        =   10
         Tag             =   "01-Job code"
         Top             =   4710
         Visible         =   0   'False
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   13
      End
      Begin VB.Label lblJob 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   96
         Top             =   4755
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblSeniority 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seniority"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   95
         Top             =   7935
         Width           =   1425
      End
      Begin VB.Label lblActBranch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Physical Branch"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   94
         Top             =   977
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblSupervisor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   93
         Top             =   4440
         Width           =   1530
      End
      Begin VB.Label lblVadim1 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
         Height          =   195
         Left            =   135
         TabIndex        =   92
         Top             =   4750
         Width           =   1545
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   90
         Top             =   285
         Width           =   555
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   89
         Top             =   631
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   88
         Top             =   1323
         Width           =   420
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   87
         Top             =   1669
         Width           =   450
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
         Left            =   135
         TabIndex        =   86
         Top             =   2361
         Width           =   1290
      End
      Begin VB.Label lblPosGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   85
         Top             =   3060
         Width           =   1395
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
         TabIndex        =   84
         Top             =   9165
         Width           =   1575
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   83
         Top             =   9555
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   82
         Top             =   9900
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Sort"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   81
         Top             =   10260
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Report"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   80
         Top             =   6180
         Width           =   1065
      End
      Begin VB.Label lblAttCodes 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Change"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   79
         Top             =   6885
         Width           =   1380
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From / To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   78
         Top             =   7590
         Width           =   1095
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
         TabIndex        =   77
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   76
         Top             =   977
         Width           =   615
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   75
         Top             =   3399
         Width           =   510
      End
      Begin VB.Label lblAdmin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Administered By"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   74
         Top             =   3745
         Width           =   1125
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   73
         Top             =   4091
         Width           =   540
      End
      Begin VB.Label lblRep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reporting Authority:"
         Height          =   195
         Left            =   135
         TabIndex        =   72
         Top             =   7245
         Width           =   1395
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   71
         Top             =   2015
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   70
         Top             =   2700
         Width           =   1200
      End
      Begin VB.Label lblShift 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7560
         TabIndex        =   69
         Top             =   4095
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblEmplStFrpmTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status From / To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   68
         Top             =   5085
         Width           =   1590
      End
      Begin VB.Label lblBand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Band"
         DataSource      =   "Data1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7500
         TabIndex        =   67
         Top             =   3750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblGrid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Category"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   66
         Top             =   3060
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblCountry 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6255
         TabIndex        =   65
         Top             =   3405
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G/L #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   64
         Top             =   4440
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prov. of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   63
         Top             =   5770
         Width           =   1455
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prov. of Residence"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   62
         Top             =   5420
         Width           =   1365
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   11400
      Top             =   10920
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
Attribute VB_Name = "frmRSALMAST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EECri As String
Dim ReportSel, HisSQL, HisSQL1
Dim WRKSQL

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkIncIndSal_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkVDesc_Click()
'Jaddy 10/24/99
fraCodeDesc.Visible = Not chkVDesc = 0
If chkVDesc = 0 Then clpCode(2).SetFocus
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
On Error Resume Next: lblCodeDesc(Index).Caption = clpCode(Index).Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Salary Master Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
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

    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "OHS", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmbBand_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)   'laura jan 06, 1997
End Sub

Private Sub comGrpLoad()
Dim vPosGroup As String
    If Not glbSyndesis Then
        vPosGroup = "Position Group"
    Else
        vPosGroup = "Position Grade"
    End If
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem lStr("Administered By")
    comGroup(0).AddItem vPosGroup '"Position Group" 'Frank 4/19/2000
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    comGroup(0).AddItem "Position Code" 'Frank Aug 6,2002 #2566
    If Not glbMulti Then comGroup(0).AddItem IIf(lStr("PShift") = "PShift", "Shift", lStr("PShift"))
    comGroup(0).AddItem lStr("Region")
    If glbWFC Then comGroup(0).AddItem "Band"
    If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14189
        If optGrouping(0).Value And optCurrHis(0).Value Then
            comGroup(0).AddItem lStr("GL #")
        End If
    End If
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem lStr("Department")
    comGroup(1).AddItem lStr("G/L")
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    comGroup(2).AddItem "Employee Name"
    comGroup(2).ListIndex = 0
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

Private Sub Cri_Vadim1()
Dim EECri As String, OneSet%, X%

If Len(clpVadim1.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_VADIM1} in ['" & Replace(clpVadim1.Text, ",", "','") & "']"

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
    'Ticket #24482 - Town of Caledon - Division from Salary (VGROUP)
    'If intIdx% = 1 Then strCd$ = IIf(glbMulti And (optCurrHis(0) Or optGrouping(1)), "HR_JOB_HISTORY.JH_ORG", "HREMP.ED_ORG")
    If intIdx% = 1 Then strCd$ = IIf(glbMulti And (optCurrHis(0) Or optGrouping(1)), "HR_JOB_HISTORY.JH_ORG", IIf(glbMulti And glbCompSerial = "S/N - 2182W" And optGrouping(0) And optCurrHis(0), "HR_JOB_HISTORY.JH_ORG", "HREMP.ED_ORG"))
    'Ticket #24482 - Town of Caledon - Division from Salary (VGROUP)
    'If intIdx% = 6 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 6 Then strCd$ = IIf(glbMulti And glbCompSerial = "S/N - 2182W" And optGrouping(0) And optCurrHis(0), "HR_JOB_HISTORY.JH_EMP", "HREMP.ED_EMP")
    If intIdx% = 7 Then strCd$ = "HRJOB.JB_GRPCD"
    If intIdx% = 8 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 9 Then strCd$ = "HREMP.ED_ADMINBY"
    'Ticket #24482 - Town of Caledon - Division from Salary (VGROUP)
    'If intIdx% = 10 Then strCd$ = "HREMP.ED_SECTION"
    If intIdx% = 10 Then strCd$ = IIf(glbMulti And glbCompSerial = "S/N - 2182W" And optGrouping(0) And optCurrHis(0), "HR_JOB_HISTORY.JH_SECTION", "HREMP.ED_SECTION")
    If intIdx% = 11 Then strCd$ = "HREMP.ED_SUPCODE"
    If intIdx% = 12 Then strCd$ = "HREMP.ED_SUBDEPT" 'Ticket #24162 Franks 10/03/2013
    
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
    'Ticket #24482 - Town of Caledon - Division from Salary (VGROUP)
    If glbCompSerial = "S/N - 2182W" And optGrouping(0) And optCurrHis(0) Then
        DivCri = "({HR_SALARY_HISTORY.SH_VGROUP} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    Else
        DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    End If
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

Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
Dim EECri As String

If Len(clpJob.Text) <= 0 Then Exit Sub
If Not optGrouping(0).Value Then
    PosCri = "HR_JOB_HISTORY.JH_JOB"
Else
    If optGrouping(0) And optCurrHis(1) Then
        PosCri = "HR_SALARY_HISTORY.SH_JOB"
    Else
        PosCri = "HR_JOB_HISTORY.JH_JOB"
    End If
End If

If Len(clpJob.Text) > 0 Then
    EECri = EECri & "({" & PosCri & "} in ['" & Replace(clpJob.Text, ",", "','") & "']) " ' AND ({HR_JOB_HISTORY.JH_CURRENT}) "
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Grid()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim GridCri As String

If Len(clpGrid.Text) <= 0 Then Exit Sub
If Not optGrouping(0).Value Then
    GridCri = "({HR_JOB_HISTORY.JH_GRID} = '" & clpGrid.Text & "')"
Else
    GridCri = "({HR_SALARY_HISTORY.SH_GRID} = '" & clpGrid.Text & "')"
End If
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & GridCri
Else
    glbstrSelCri = GridCri
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

Private Sub Cri_ActingPosition()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
    
    xTemp = False
    EECri = ""
    If chkActing.Value = 0 Then
        Exit Sub
    End If
    If (optCurrHis(1) Or optCurrHis(2)) And optGrouping(0) Then
        Exit Sub
    End If

    EECri = EECri & "{HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES' "
    xTemp = True

    If Len(EECri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    End If

End Sub

Private Sub Cri_RepAuth()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
    
    xTemp = False
    EECri = ""
    If optGrouping(2) Then
        Exit Sub
    End If
    
    If (optCurrHis(1) Or optCurrHis(2)) And optGrouping(0) Then
        Exit Sub
    End If
    
    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(3).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
        End If
        xTemp = True
    End If
    
    If Len(EECri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    End If

End Sub

Private Sub Cri_AsDates()
Dim TempCri As String
Dim TempCri01 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%, strCriNam$

If optGrouping(2) Then
    Exit Sub
End If
    
If Not optGrouping(0) And Not optGrouping(1) Then optGrouping(0) = True
If optGrouping(0).Value = True Then
    strCriNam$ = "HR_SALARY_HISTORY.SH_EDATE"
    GoTo DoCrit
End If

If optGrouping(1).Value = True Then
    strCriNam$ = "HR_JOB_HISTORY.JH_SDATE"
    GoTo DoCrit
End If
Exit Sub  ' if none of these set then exit routine

DoCrit:

If Len(dlpDateAs.Text) > 0 Then
    TempCri = "({" & strCriNam$ & "} "
    TempCri01 = "((" & strCriNam$ & ") " & " <= " & Date_SQL(dlpDateAs.Text) & ")"
    
    WRKSQL = WRKSQL & " AND " & TempCri01
    TempCri = TempCri & " <= "
    dtYYY% = Year(dlpDateAs.Text)
    dtMM% = month(dlpDateAs.Text)
    dtDD% = Day(dlpDateAs.Text)
    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%, strCriNam$

If Not optGrouping(0) And Not optGrouping(1) And Not optGrouping(2) Then optGrouping(0) = True
If optGrouping(0).Value = True Then
    strCriNam$ = "HR_SALARY_HISTORY.SH_EDATE"
    GoTo DoCrit
End If

If optGrouping(1).Value = True Or optGrouping(2).Value = True Then
    strCriNam$ = "HR_JOB_HISTORY.JH_SDATE"
    GoTo DoCrit
End If
Exit Sub  ' if none of these set then exit routine

DoCrit:
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({" & strCriNam$ & "} "
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
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({" & strCriNam$ & "} "
        If X% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub

'Ticket #24482 - Town of Caledon - Division from Salary (VGROUP)
'If glbMulti And (optCurrHis(0) Or optGrouping(1)) Then
If (glbMulti And (optCurrHis(0) Or optGrouping(1))) Or (glbCompSerial = "S/N - 2182W" And optGrouping(0) And optCurrHis(0)) Then
    EECri = "{HR_JOB_HISTORY.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
Else
    EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Reason()
Dim EECri As String, OneSet%, X%
Dim SR1, SR2, SR3, SRR

OneSet% = False

If optGrouping(2) Then
    Exit Sub
End If

For X% = 2 To 5
    If Len(clpCode(X%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then Exit Sub

For X% = 2 To 5
    If Len(clpCode(X%).Text) > 0 Then
        SRR = SRR & "'" & clpCode(X%).Text & "' "
        If OneSet% > 1 Then SRR = SRR & ","
        OneSet% = OneSet% - 1
    End If
Next X%

If optGrouping(1) Then           'laura 03/09/98
    SR1 = "{HR_JOB_HISTORY.JH_JREASON} IN [" & SRR & "]"
    EECri = "(" & SR1 & ") "
Else

SR1 = "{HR_SALARY_HISTORY.SH_SREAS1} IN [" & SRR & "]"
SR2 = "{HR_SALARY_HISTORY.SH_SREAS2} IN [" & SRR & "]"
SR3 = "{HR_SALARY_HISTORY.SH_SREAS3} IN [" & SRR & "]"

EECri = "((" & SR1 & ") OR (" & SR2 & ") OR (" & SR3 & "))"
'EECri = "(" & SR1 & ") "

End If

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
Dim X%, LocstrSelCri

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div
Call Cri_Code(0)
Call Cri_Code(1)
Call Cri_Position
Call Cri_Grid
For X% = 6 To 10
    Call Cri_Code(X%)
Next
Call Cri_PT
Call Cri_Shift
Call Cri_CountryOfEmployment 'Ticket #13621 Frank 09/07/2007
Call Cri_GLNO 'Ticket #14189
If glbWFC Then Call Cri_Band
Call Cri_EE
Call Cri_Reason
Call Cri_FTDates
Call Cri_EmpStatFTDates
Call Cri_RepAuth
Call Cri_ActingPosition
Call Cri_ProvResidence
Call Cri_ProvEmployment
Call Cri_Code(11) 'Ticket #20600 Franks 09/22/2011
Call Cri_Vadim1 'Ticket #20600 Franks 09/22/2011
Call Cri_Code(12) 'Ticket #24162 Franks 10/03/2013
Call Cri_SenDates 'Ticket #24162 Franks 10/03/2013

If glbWFC Then 'Ticket #27531 Franks 09/15/2015
    Call Cri_JobMaster
End If
    
WRKSQL = Replace(Replace(glbstrSelCri, "{", "("), "}", ")")
WRKSQL = Replace(Replace(WRKSQL, "[", "("), "]", ")")

Call Cri_AsDates
 
If optCurrHis(2) Then
    If optGrouping(0) Then Call SalWRK
    If optGrouping(1) Then Call JobWRK
End If

If optGrouping(0) Then
    If optCurrHis(0) Then
        If glbLinamar Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rlSalsum.rpt"
        ElseIf glbCompSerial = "S/N - 2334W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2334_1.rpt"
            Me.vbxCrystal.Formulas(4) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
        ElseIf glbMulti Then
            If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14189
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmSalsumdnssab.rpt"
                'Ticket #15725
                Me.vbxCrystal.Formulas(7) = "lblGL='" & lStr("G/L#") & "'"
            Else
                Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
                'Ticket #24482 - Town of Caledon.
                If glbCompSerial = "S/N - 2182W" Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2182_rmSalsum.rpt"
                Else
                    'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
                    Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
                
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmSalsum.rpt"
                End If
            End If
        ElseIf glbCompSerial = "S/N - 2335W" Then 'Ticket #23821 - Mitchell Plastics
            Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
            Me.vbxCrystal.Formulas(11) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
            Me.vbxCrystal.Formulas(12) = "lblUnion='" & lStr("Union") & "'"
            Me.vbxCrystal.Formulas(13) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(14) = "lblLocation='" & lStr("Location") & "'"
            
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2335_rzSalsum.rpt"
        ElseIf glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls Ticket #27681 Franks 12/10/2015
            Call NiagaraFallsDefUpt
            Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2276_rzSalsum.rpt"
        Else
            Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
            
            'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
            Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
            
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzSalsum.rpt"
        End If
        
        If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #17324
            Me.vbxCrystal.Formulas(8) = "lblPosGroup='" & lStr("Position Group") & "'"
        End If
        
        Me.vbxCrystal.WindowTitle = "Current Salary Summary Report"
    ElseIf optCurrHis(1) Then
        If glbCompSerial = "S/N - 2334W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2334_2.rpt"
            Me.vbxCrystal.Formulas(4) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
        Else
            Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
            
            'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
            Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
            
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzSalHis.rpt"
        End If
        
        If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #17324
            Me.vbxCrystal.Formulas(8) = "lblPosGroup='" & lStr("Position Group") & "'"
        End If
        
        Me.vbxCrystal.WindowTitle = "Salary History Report"
    ElseIf optCurrHis(2) Then
        If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #17324
            Me.vbxCrystal.Formulas(8) = "lblPosGroup='" & lStr("Position Group") & "'"
        End If
        Me.vbxCrystal.Formulas(10) = "showCompaRatio = " & IIf(gsCompaRatio = 0, False, True) & " "
        
        'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
        Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
        
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzSalash.rpt"
        Me.vbxCrystal.WindowTitle = "Salary Summary Report"
    End If
Else
    If optCurrHis(0) Then
        If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls Ticket #27681 Franks 12/10/2015
            Call NiagaraFallsDefUpt
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2276_rzPosSum.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzPosSum.rpt"
        End If
        
        Me.vbxCrystal.WindowTitle = "Current Position Summary Report"
        
        If glbCompSerial = "S/N - 2359W" Then 'Ticket# 7644
            If Not (glbSQL Or glbOracle) Then
                Me.vbxCrystal.Formulas(3) = "lblUnion='" & lStr("Union") & "'"
            End If
        End If
        
        If glbMulti And (Not glbSQL And Not glbOracle) Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmPossum.rpt"
        End If
        
        If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls Ticket #27681 Franks 12/10/2015
            'don't use this
        Else
            'Ticket #23537 - Release 8.0
            Me.vbxCrystal.Formulas(11) = "lblDHrs='" & lStr("Hours/Day") & "'"
            Me.vbxCrystal.Formulas(12) = "lblWHrs='" & lStr("Hours/Week") & "'"
            Me.vbxCrystal.Formulas(13) = "lblPHrs='" & lStr("Hours/Pay Period") & "'"
            Me.vbxCrystal.Formulas(14) = "lblFTEHrsYr='" & lStr("FTE Hours/Year") & "'"
        
            'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
            'Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
        End If
        
    ElseIf optCurrHis(1) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzPosHis.rpt"
        Me.vbxCrystal.WindowTitle = "Position History Report"
        
        'Ticket #23537 - Release 8.0
        Me.vbxCrystal.Formulas(11) = "lblDHrs='" & lStr("Hours/Day") & "'"
        Me.vbxCrystal.Formulas(12) = "lblWHrs='" & lStr("Hours/Week") & "'"
        Me.vbxCrystal.Formulas(13) = "lblPHrs='" & lStr("Hours/Pay Period") & "'"
        Me.vbxCrystal.Formulas(14) = "lblFTEHrsYr='" & lStr("FTE Hours/Year") & "'"
        
        'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
        'Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
        
    ElseIf optCurrHis(2) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzPosash.rpt"
        Me.vbxCrystal.WindowTitle = "Position Summary Report"
    
        'Ticket #23537 - Release 8.0
        Me.vbxCrystal.Formulas(11) = "lblDHrs='" & lStr("Hours/Day") & "'"
        Me.vbxCrystal.Formulas(12) = "lblWHrs='" & lStr("Hours/Week") & "'"
        Me.vbxCrystal.Formulas(13) = "lblPHrs='" & lStr("Hours/Pay Period") & "'"
        Me.vbxCrystal.Formulas(14) = "lblFTEHrsYr='" & lStr("FTE Hours/Year") & "'"
    
        'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
        'Me.vbxCrystal.Formulas(15) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
    End If
End If

If optGrouping(2) Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZPosBackup.rpt"
End If

X% = Cri_Sorts()
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    If Not optCurrHis(2) Then  'As of Date
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        If optGrouping(0) Then 'Salary
            For X% = 0 To 14
                If X% = 3 Then
                    Me.vbxCrystal.DataFiles(X%) = glbIHRDBW
                Else
                    Me.vbxCrystal.DataFiles(X%) = glbIHRDB
                End If
            Next X%
        End If
        If optGrouping(1) Then 'Position
            For X% = 0 To 15
                If X% = 4 Then
                    Me.vbxCrystal.DataFiles(X%) = glbIHRDBW
                Else
                    Me.vbxCrystal.DataFiles(X%) = glbIHRDB
                End If
            Next X%
        End If
    End If
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'CONP'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'CONP'"
            End If
        End If
        If chkExclRET Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'RET'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'RET'"
            End If
        End If
    End If
End If

'If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
'    Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"    'Hemu -EXE
'ElseIf glbNoNONE Then
'    Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
'ElseIf glbNoEXEC Then    'Hemu -EXE
'    Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"    'Hemu -EXE
'Else
'    Me.vbxCrystal.SelectionFormula = glbstrSelCri
'End If
If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
    LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"    'Hemu -EXE
ElseIf glbNoNONE Then
    LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
ElseIf glbNoEXEC Then    'Hemu -EXE
    LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"    'Hemu -EXE
Else
    LocstrSelCri = glbstrSelCri
End If

If glbWFC Then
    If optGrouping(0) Then 'For Salary Only (Band Security)
        If gSec_WFC_Band_Security Then
            LocstrSelCri = LocstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
        End If
    End If
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    'Limit the user has access to these employees salary information
    If Len(glbNoAccessGrp) <> 0 And optGrouping(0) Then
        If Len(LocstrSelCri) > 0 Then
            LocstrSelCri = LocstrSelCri & " AND (" & glbNoAccessGrp & ") "
        Else
            LocstrSelCri = LocstrSelCri & " (" & glbNoAccessGrp & ") "
        End If
        LocstrSelCri = Replace(Replace(LocstrSelCri, "HREMP.ED_", "ED_"), "ED_", "HREMP.ED_")
        LocstrSelCri = Replace(LocstrSelCri, "(HREMP.", "({HREMP.")
        LocstrSelCri = Replace(LocstrSelCri, " HREMP.", " {HREMP.")
        LocstrSelCri = Replace(LocstrSelCri, " =", "} =")
    End If
End If

Me.vbxCrystal.SelectionFormula = LocstrSelCri

If optCurrHis(2) And Not optGrouping(2) Then  'As of Date
    If optGrouping(0) Then
        LocstrSelCri = "{HR_SALARY_HISTORY.WRKEMP}= '" & glbUserID & "'"
        If glbWFC Then
            If gSec_WFC_Band_Security Then
                LocstrSelCri = LocstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
            End If
        End If
        Me.vbxCrystal.SelectionFormula = LocstrSelCri
    End If
    If optGrouping(1) Then
        LocstrSelCri = "{HR_JOB_HISTORY.WRKEMP}= '" & glbUserID & "'"
        Me.vbxCrystal.SelectionFormula = LocstrSelCri
    End If
    Me.vbxCrystal.Formulas(5) = "AsDate = 'As of Date = " & CVDate(dlpDateAs.Text) & "' "
End If

If glbWFC Then
    Me.vbxCrystal.Formulas(6) = "WFCflag = True "
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
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub JobWRK()
Dim rsJOB As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance
Dim iFieldCounter, FldName, xNum, xNumT
Dim EmpFlag As Boolean, xEmpNo
Screen.MousePointer = HOURGLASS
SQLQ = "DELETE FROM HRJOBWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

SQLQ = "SELECT HR_JOB_HISTORY.* "
If glbOracle Then
    SQLQ = SQLQ & " FROM HR_JOB_HISTORY, HREMP ,HRJOB WHERE HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB (+) = HRJOB.JB_CODE "
    SQLQ = SQLQ & " AND " & WRKSQL
Else
    SQLQ = SQLQ & " FROM (HR_JOB_HISTORY INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR) "
    SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " WHERE " & WRKSQL
End If

'Frank Ticket# 6795 Missed Security of -NON and -EXE
If glbNoNONE And glbNoEXEC Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'NONE' AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'EXEC' ) "
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'NONE' ) "
ElseIf glbNoEXEC Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'EXEC' ) "
End If

'Ticket #21894 - Order by Current Position first and then by Start Date. In this way picking up Current
'Position if it is older than the non current position but still meeting the As Of Date criteria.
SQLQ = SQLQ & " ORDER BY JH_EMPNBR,JH_SDATE DESC "
'If Jerry says 'go ahead' then turning this sort ON and above one OFF.
'SQLQ = SQLQ & " ORDER BY JH_EMPNBR,JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"


rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsJOB.EOF Then
    rsJOB.Close
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
rsJOB.MoveLast
rsJOB.MoveFirst
xNumT = rsJOB.RecordCount
xNum = 0
rsWRK.Open "HRJOBWRK", IIf(Not glbSQL And Not glbOracle, gdbAdoIhr001W, gdbAdoIhr001), adOpenKeyset, adLockOptimistic, adCmdTableDirect
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(1).Caption = "Please Wait"
EmpFlag = True: xEmpNo = 0
Do While Not rsJOB.EOF
    MDIMain.panHelp(0).FloodPercent = (xNum / xNumT) * 100
    xNum = xNum + 1
    If xEmpNo <> rsJOB("JH_EMPNBR") Then
        xEmpNo = rsJOB("JH_EMPNBR")
        EmpFlag = True
    End If
    If rsJOB("JH_SDATE") > CVDate(dlpDateAs.Text) Then
        GoTo Next_Rec
    End If
    If Not EmpFlag Then
        GoTo Next_Rec
    End If

    EmpFlag = False

    rsWRK.AddNew
    For iFieldCounter = 0 To rsWRK.Fields.count - 1
        FldName = rsWRK.Fields(iFieldCounter).name
        If FldName <> "JH_ID" And FldName <> "WRKEMP" Then
            rsWRK(FldName) = rsJOB(FldName)
        End If
        If FldName = "WRKEMP" Then
            rsWRK(FldName) = glbUserID
        End If
    Next iFieldCounter
    rsWRK.Update
Next_Rec:
    rsJOB.MoveNext
Loop
rsWRK.Close
rsJOB.Close
If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT
End Sub

Private Sub SalWRK()
Dim rsSal As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance
Dim iFieldCounter, FldName, xNum, xNumT
Dim EmpFlag As Boolean, xEmpNo
Screen.MousePointer = HOURGLASS

SQLQ = "DELETE FROM HRSALWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

SQLQ = "SELECT HR_SALARY_HISTORY.* "
If glbOracle Then
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HREMP, HRJOB WHERE  HR_SALARY_HISTORY.SH_EMPNBR = HREMP.ED_EMPNBR  "
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " AND  " & WRKSQL
Else
    WRKSQL = Replace(WRKSQL, "HR_JOB_HISTORY.JH_JOB", "HRJOB.JB_CODE") 'Franks May 12,2003
    SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HREMP ON HR_SALARY_HISTORY.SH_EMPNBR = HREMP.ED_EMPNBR) "
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " WHERE " & WRKSQL
End If

'Ticket #24610 - If RAs are entered then it was not retrieving records based on that.
Dim xTemp As Boolean
xTemp = False
If Len(Trim(elpRept(0).Text)) > 0 Then
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR = SH_EMPNBR AND JH_CURRENT <> 0 AND JH_REPTAU = " & Trim(elpRept(0).Text)
    xTemp = True
End If
If Len(Trim(elpRept(1).Text)) > 0 Then
    If xTemp = True Then
        SQLQ = SQLQ & " AND JH_REPTAU2 = " & Trim(elpRept(1).Text)
    Else
        SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR = SH_EMPNBR AND JH_CURRENT <> 0 AND JH_REPTAU2 = " & Trim(elpRept(1).Text)
    End If
    xTemp = True
End If
If Len(Trim(elpRept(2).Text)) > 0 Then
    If xTemp Then
        SQLQ = SQLQ & " AND JH_REPTAU3 = " & Trim(elpRept(2).Text)
    Else
        SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR = SH_EMPNBR AND JH_CURRENT <> 0 AND JH_REPTAU3 = " & Trim(elpRept(2).Text)
    End If
    xTemp = True
End If
If Len(Trim(elpRept(3).Text)) > 0 Then
    If xTemp Then
        SQLQ = SQLQ & " AND JH_REPTAU4 = " & Trim(elpRept(3).Text)
    Else
        SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR = SH_EMPNBR AND JH_CURRENT <> 0 AND JH_REPTAU4 = " & Trim(elpRept(3).Text)
    End If
    xTemp = True
End If
If xTemp = True Then SQLQ = SQLQ & ")"

'Frank Ticket# 6795 Missed Security of -NON and -EXE
If glbNoNONE And glbNoEXEC Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'NONE' AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'EXEC' ) "
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'NONE' ) "
ElseIf glbNoEXEC Then
    SQLQ = SQLQ & " AND (HREMP.ED_ORG IS NULL OR HREMP.ED_ORG<>'EXEC' ) "
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    'Limit the user has access to these employee's salary information
    If Len(glbNoAccessGrp) <> 0 Then
        SQLQ = SQLQ & " AND (" & glbNoAccessGrp & ") "
    End If
End If

SQLQ = SQLQ & " ORDER BY SH_EMPNBR,SH_EDATE DESC "

rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsSal.EOF Then
    rsSal.Close
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
xNumT = rsSal.RecordCount
xNum = 0
rsWRK.Open "HRSALWRK", IIf(Not glbSQL And Not glbOracle, gdbAdoIhr001W, gdbAdoIhr001), adOpenKeyset, adLockOptimistic, adCmdTableDirect
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(1).Caption = "Please Wait"
EmpFlag = True: xEmpNo = 0
Do While Not rsSal.EOF
    MDIMain.panHelp(0).FloodPercent = (xNum / xNumT) * 100
    xNum = xNum + 1
    If xEmpNo <> rsSal("SH_EMPNBR") Then
        xEmpNo = rsSal("SH_EMPNBR")
        EmpFlag = True
    End If
    If rsSal("SH_EDATE") > CVDate(dlpDateAs.Text) Then
        GoTo Next_Rec
    End If
    If Not EmpFlag Then
        GoTo Next_Rec
    End If

    EmpFlag = False

    rsWRK.AddNew
    For iFieldCounter = 0 To rsWRK.Fields.count - 1
        FldName = rsWRK.Fields(iFieldCounter).name
        If FldName <> "SH_ID" And FldName <> "WRKEMP" Then
            rsWRK(FldName) = rsSal(FldName)
        End If
        If FldName = "WRKEMP" Then
            rsWRK(FldName) = glbUserID
        End If
    Next iFieldCounter
    rsWRK.Update
Next_Rec:
    rsSal.MoveNext
Loop
rsWRK.Close
rsSal.Close
If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$, DOA
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3

z% = 0
Cri_Sorts = 0

grpField$ = getEGroup(comGroup(0).Text)

'Frank Ticket# 6587 Add "Band" to Grouping0 for WFC
If comGroup(0).Text = "Band" Then
    If optGrouping(0) Then
        grpField$ = "{HR_SALARY_HISTORY.SH_BAND}"
    Else
        grpField$ = "{HRJOB.JB_BAND}"
    End If
End If

If grpField$ = "(none)" Then grpField$ = "{HRPARCO.PC_NAME}"
If comGroup(0) = lStr("Union") Then
    If glbMulti And (optCurrHis(0) Or optGrouping(1)) Then
        grpField$ = "{tblUnion.TB_DESC}"    '"{HR_JOB_HISTORY.JH_ORG}"
    Else
        grpField$ = "{tblUnion.TB_DESC}"    '"{HREMP.ED_ORG}"
    End If
End If
'' first set primary grouping ------------------------- **
'GrpIdx% = comGroup(0).ListIndex
'If Not glbLinamar Then ' Frank May 2, 2001
'    Select Case GrpIdx%
'        Case 0: grpField$ = "{HR_DIVISION.Division_Name}"
'        Case 1: grpField$ = "{HRDEPT.DF_NAME}"
'        Case 2: grpField$ = "{HRTABL.TB_DESC}"
'        Case 3
'            If glbMulti And (optCurrHis(0) Or optGrouping(1)) Then
'                grpField$ = "{HR_JOB_HISTORY.JH_ORG}"
'            Else
'                grpField$ = "{HREMP.ED_ORG}"
'            End If
'        Case 4: grpField$ = "{TblAdminby.TB_DESC}"
'        Case 5: grpField$ = "{tblPosGroup.TB_DESC}" 'Frank 4/19/2000
'        Case 6: grpField$ = "{tblSec.TB_DESC}"
'        Case 7: grpField$ = "{@EFullName}"
'        Case 8: grpField$ = "{HRJOB.JB_DESCR}" 'Frank Aug 6,2002 #2566
'        Case 9: grpField$ = "{HREMP.ED_SHIFT}"
'        Case 10: grpField$ = "{HRPARCO.PC_NAME}"
'    End Select
'Else
'    Select Case GrpIdx%
'        Case 0: grpField$ = "{HR_DIVISION.Division_Name}"
'        Case 1: grpField$ = "{HRDEPT.DF_NAME}"
'        Case 2: grpField$ = "{HRTABL.TB_DESC}"
'        Case 3
'            If glbMulti And (optCurrHis(0) Or optGrouping(1)) Then
'                grpField$ = "{HR_JOB_HISTORY.JH_ORG}"
'            Else
'                grpField$ = "{HREMP.ED_ORG}"
'            End If
'        Case 4: grpField$ = "{TblAdminby.TB_DESC}"
'        Case 5: grpField$ = "{tblPosGroup.TB_DESC}" 'Frank 4/19/2000
'        Case 6: grpField$ = "{tblSec.TB_DESC}"
'        Case 7: grpField$ = "{@EFullName}"
'        Case 8: grpField$ = "{@EMPTYPE}"
'        Case 9: grpField$ = "{tblRegion.TB_DESC}"
'        Case 10: grpField$ = "{LN_HOMES.TB_DESC}"
'        Case 11: grpField$ = "{HRJOB.JB_DESCR}" 'Frank Aug 6,2002 #2566
'        Case 12: grpField$ = "{HREMP.ED_SHIFT}"
'        Case 13: grpField$ = "{HRPARCO.PC_NAME}"
'    End Select
'
'End If

SavGrp1 = grpField$
If Not (grpField$ = "{HRPARCO.PC_NAME}") Then 'GrpIdx% < 8 Then
    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
Else
    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
End If
Me.vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"

' SECOND primary grouping ------------------------- **
GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{HRDEPT.DF_NAME}"
    Case 1: grpField$ = "{HRGL.GL_DESCR}"
    Case 2: grpField$ = "{HRPARCO.PC_NAME}" '"(none)"
End Select
Me.vbxCrystal.GroupCondition(1) = "GROUP2;" & grpField$ & ";ANYCHANGE;A"
If grpField$ = "{HRPARCO.PC_NAME}" Then
    Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
Else
    Me.vbxCrystal.SectionFormat(1) = "GF2;T;X;X;X;X;X;X"
End If
' END ---------------------------------------------- ***

' Final Sort primary grouping ------------------------- **
GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "(none)"
End Select
Me.vbxCrystal.GroupCondition(2) = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
' END ---------------------------------------------- ***

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%, intDSet%, I

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
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 10
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

'If Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
'    MsgBox lStr("If Position Code Entered - it must be known")
'    clpJob.SetFocus
'    Exit Function
'End If

'Add by Franks Mar 4,2002 Begin
If optCurrHis(0).Value Then
    If IsDate(dlpDateAs.Text) Then
        MsgBox "If Current selected, As of Date must be blank"
        dlpDateAs.SetFocus
        Exit Function
    End If
End If

If optCurrHis(1).Value Then
    If IsDate(dlpDateAs.Text) Then
        MsgBox "If History selected, As of Date must be blank"
        dlpDateAs.SetFocus
        Exit Function
    End If
End If

If optCurrHis(2).Value Then
    If Not IsDate(dlpDateAs.Text) Then
        MsgBox "If As of Date selected, you must enter a valid date"
        dlpDateAs.SetFocus
        Exit Function
    End If
    If IsDate(dlpDateRange(0).Text) Then
        MsgBox "If As of Date selected, From Date must be blank"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
    If IsDate(dlpDateRange(1).Text) Then
        MsgBox "If As of Date selected, To Date must be blank"
        dlpDateRange(1).SetFocus
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

'Add by Franks Mar 4,2002 End
intDSet% = False
For X% = 0 To 3
    If Len(dlpDateRange(X%).Text) > 0 Then
       If Not IsDate(dlpDateRange(X%).Text) Then
           MsgBox "Not a valid date"
           dlpDateRange(X%).Text = ""
           dlpDateRange(X%).SetFocus
           Exit Function
       End If
       intDSet% = True
    End If
Next X%

For I = 0 To 3
    If elpRept(I).Caption = "Enter Valid Employee #" Then
        MsgBox "If Reporting Authority Entered - they must exist"
        elpRept(I).SetFocus
        Exit Function
    End If
Next

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

If IsDate(dlpDateRange(3).Text) And IsDate(dlpDateRange(2).Text) Then
    If DaysBetween(dlpDateRange(2).Text, dlpDateRange(3).Text) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(2).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE

'Hamilton CAS - Ticket #13398
If glbCompSerial = "S/N - 2257W" Then
    If chkActing.Visible = True Then
        chkActing.Caption = "Red Circled"
    End If
End If

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

'Ticket #16228 - Oxford
If glbCompSerial = "S/N - 2259W" Then
    glbMulti = True
End If

'Burlington Tech - Ticket #13235
If Not glbCompSerial = "S/N - 2351W" Then
    optGrouping(2).Visible = False
End If

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
    lblShift.Caption = lStr("PShift")
    If lblShift.Caption = "PShift" Then lblShift.Caption = "Shift"
End If

'Frank Nov 4,2002
If glbWFC Then
    'cmbBand.Visible = True
    clpCode(13).Visible = True 'Ticket #26507 Franks 01/14/2015
    lblBand.Visible = True
    'cmbBand.AddItem ""
    'Band_AddItem Me
Else
    cmbBand.Visible = False
    lblBand.Visible = False
End If

If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If

'Ticket #24122 - Label Master
chkActing.Caption = lStr(chkActing.Caption)

If (optCurrHis(1) Or optCurrHis(2)) And optGrouping(0) Then
    chkActing.Visible = False
Else
    chkActing.Visible = True
End If

If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14189
    lblTitle(1).Visible = True
    clpGLNum.Visible = True
End If

'Frank Nov 4,2002
Call comGrpLoad
Call setRptCaption(Me)

'Ticket #20600 Franks 09/22/2011
lblSupervisor.Caption = lStr("Supervisor Code")
lblVadim1.Caption = lStr("Vadim Field 1")
lblSeniority.Caption = lStr("Seniority")

lblTitle(1).Caption = lStr(lblTitle(1).Caption)
If glbLinamar Then clpCode(8).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(8).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbSyndesis Then
    lblPosGroup.Caption = "Position Grade"
    clpCode(7).Tag = "00-Enter Position Grade"
End If
If glbCompSerial = "S/N - 2172W" Then 'Lanark
    lblPosGroup.Caption = "Salary Level"
    clpCode(7).TABLTitle = "00-Enter Salary Level"
End If

'Hamilton CAS - Ticket #13398
If glbCompSerial = "S/N - 2257W" Then
    If chkActing.Visible = True Then
        chkActing.Caption = "Red Circled"
    End If
End If

Call addCountryItems 'Frank 09/07/2007 Ticket #13621

If glbSamuel Then 'Ticket #24162 Franks 09/20/2013
    Call SamuelScreenSetup
End If

If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0
Call INI_Controls(Me)

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob.TransDiv = glbWFCUserSecList
    
    'Ticket #27531 Franks 09/14/2015
    Call WFCScreenSetup
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    chkIncIndSal.Left = chkExclCONP.Left
    chkIncIndSal.Visible = True
Else
    chkIncIndSal.Value = False
    chkIncIndSal.Visible = False
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
Dim c As Long

On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 10500 Then
        scrControl.Value = 0
        
        pcSalary.Top = 120
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 1000
        
        scrControl.Max = 2200
        
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Salary Master Report", "edit/Add")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optCurrHis_Click(Index As Integer, Value As Integer)
'Jaddy 10/24/99
    If (optCurrHis(1) Or optCurrHis(2)) And optGrouping(0) Then
        chkActing.Visible = False
    Else
        chkActing.Visible = True
    End If
    If glbCompSerial = "S/N - 2388W" Then 'DNSSAB
        Call comGrpLoad 'Ticket #14189
    End If
End Sub

Private Sub optCurrHis_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)    'laura jan 06, 1997
End Sub

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
    If (optCurrHis(1) Or optCurrHis(2)) And optGrouping(0) Then
        chkActing.Visible = False
    Else
        chkActing.Visible = True
    End If
    Call BackupPosScreen(Index)
    
    If glbCompSerial = "S/N - 2388W" Then 'DNSSAB
        Call comGrpLoad 'Ticket #14189
    End If

    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" And optGrouping(0) Then
        chkIncIndSal.Visible = True
    Else
        chkIncIndSal.Visible = False
    End If
    
End Sub

Private Sub BackupPosScreen(xNo As Integer)
Dim xFlag As Boolean
    If xNo = 2 Then
        xFlag = False
        chkVDesc.Value = False
        chkActing.Value = False
        clpCode(2).Text = ""
        clpCode(3).Text = ""
        clpCode(4).Text = ""
        clpCode(5).Text = ""
        elpRept(0).Text = ""
        elpRept(1).Text = ""
        elpRept(2).Text = ""
        elpRept(3).Text = ""
    Else
        xFlag = True
    End If
    lblAttCodes.Enabled = xFlag
    clpCode(2).Enabled = xFlag
    clpCode(3).Enabled = xFlag
    clpCode(4).Enabled = xFlag
    clpCode(5).Enabled = xFlag
    chkVDesc.Enabled = xFlag
    lblRep.Enabled = xFlag
    elpRept(0).Enabled = xFlag
    elpRept(1).Enabled = xFlag
    elpRept(2).Enabled = xFlag
    elpRept(3).Enabled = xFlag
    chkActing.Enabled = xFlag
    
End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)   'laura jan 06, 1997
End Sub

Private Sub scrControl_Change()
pcSalary.Top = 120 - scrControl.Value
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

Private Sub Cri_EmpStatFTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, X%
Dim FromDate, toDate, SQLQ
Dim rsHRPARCO As New ADODB.Recordset

If optGrouping(2) Then
    Exit Sub
End If
    
If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
    TempCri = "({HREMP.ED_SFDATE} "
    dtYYY% = Year(dlpDateRange(2).Text)
    dtMM% = month(dlpDateRange(2).Text)
    dtDD% = Day(dlpDateRange(2).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(3).Text)
    dtMM% = month(dlpDateRange(3).Text)
    dtDD% = Day(dlpDateRange(3).Text)
    TempCri = TempCri & " ({HREMP.ED_STDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 Then
        TempCri = "({HREMP.ED_STDATE} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    If Len(dlpDateRange(3).Text) > 0 Then
        TempCri = TempCri & "({HREMP.ED_SFDATE}  "
        TempCri = TempCri & " <= "
        dtYYY% = Year(dlpDateRange(3).Text)
        dtMM% = month(dlpDateRange(3).Text)
        dtDD% = Day(dlpDateRange(3).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Else
    GoTo Cri_FTDatst
End If

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

Private Sub Cri_Band()
Dim EECri As String, OneSet%, X%

'If Len(cmbBand.Text) < 1 Then Exit Sub
'
'If optGrouping(0).Value Then
'    EECri = "{HR_SALARY_HISTORY.SH_BAND}= '" & cmbBand.Text & "'"
'Else
'    EECri = "{HRJOB.JB_BAND}= '" & cmbBand.Text & "'"
'End If
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True

'Ticket #26507 Franks 01/14/2015 - begin
If Len(clpCode(13).Text) <= 0 Then Exit Sub

If optGrouping(0).Value Then
    EECri = "{HR_SALARY_HISTORY.SH_BAND} in ['" & Replace(clpCode(13).Text, ",", "','") & "'] "
Else
    EECri = "{HRJOB.JB_BAND} in ['" & Replace(clpCode(13).Text, ",", "','") & "'] "
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
'Ticket #26507 Franks 01/14/2015 - end

End Sub

Private Sub addCountryItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountryOfEmp.AddItem ctylist
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

Private Sub SamuelScreenSetup() 'Ticket #24162 Franks 10/03/2013
    lblActBranch.Visible = True
    clpCode(12).Visible = True
    'lblSeniority.Visible = True
    'lblToDate.Visible = True
    'dlpSenDateRange(0).Visible = True
    'dlpSenDateRange(1).Visible = True
End Sub

Private Sub Cri_SenDates() 'Ticket #24162 Franks 10/03/2013
Dim TempCri As String
Dim TempCri2 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpSenDateRange(0).Text) = 0 And Len(dlpSenDateRange(1).Text) = 0 Then
    Exit Sub
End If

If Len(dlpSenDateRange(0).Text) > 0 And Len(dlpSenDateRange(1).Text) > 0 Then
    TempCri = "({HREMP.ED_SENDTE}) "
    TempCri2 = "({HREMP.ED_SENDTE}) "

    dtYYY% = Year(dlpSenDateRange(0).Text)
    dtMM% = month(dlpSenDateRange(0).Text)
    dtDD% = Day(dlpSenDateRange(0).Text)
    'TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpSenDateRange(1).Text)
    dtMM% = month(dlpSenDateRange(1).Text)
    dtDD% = Day(dlpSenDateRange(1).Text)
    TempCri2 = TempCri2 & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " AND " & TempCri2
    GoTo Cri_SenDatst
End If

For X% = 0 To 1
    TempCri = "({HREMP.ED_SENDTE}) "
    TempCri2 = "({HREMP.ED_SENDTE}) "
        
    If Len(dlpSenDateRange(0).Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpSenDateRange(0).Text)
        dtMM% = month(dlpSenDateRange(0).Text)
        dtDD% = Day(dlpSenDateRange(0).Text)
    End If
    If Len(dlpSenDateRange(1).Text) > 0 Then
        TempCri = TempCri2 & " <= "
        dtYYY% = Year(dlpSenDateRange(1).Text)
        dtMM% = month(dlpSenDateRange(1).Text)
        dtDD% = Day(dlpSenDateRange(1).Text)
    End If

    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    GoTo Cri_SenDatst

Next X%



Cri_SenDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub WFCScreenSetup() 'Ticket #27531 Franks 09/14/2015
    lblJob.Top = lblGrid.Top
    clpJobMaster.Top = clpGrid.Top
    lblJob.Visible = True
    clpJobMaster.Visible = True
End Sub

Private Sub Cri_JobMaster() 'Ticket #27531 Franks 09/15/2015
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String

If Len(clpJobMaster.Text) <= 0 Then Exit Sub

PosCri = "({HRJOB.JB_JOBCODE} = '" & clpJobMaster.Text & "')"

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If
End Sub

Private Sub NiagaraFallsDefUpt() 'Ticket #27531 Franks 09/15/2015
Dim SQLQ As String
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ESTATUS_TABL = 'EPST' WHERE JH_ESTATUS_TABL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_ESTATUS_TABL = 'EPST' WHERE JH_ESTATUS_TABL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_POSITION_CONTROL = 'NO' WHERE JH_POSITION_CONTROL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_POSITION_CONTROL = 'NO' WHERE JH_POSITION_CONTROL IS NULL"
    gdbAdoIhr001.Execute SQLQ
    
End Sub

