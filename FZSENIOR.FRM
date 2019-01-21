VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRSeniority 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Seniority Report"
   ClientHeight    =   10485
   ClientLeft      =   195
   ClientTop       =   810
   ClientWidth     =   12270
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
   ScaleHeight     =   10485
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   9075
      LargeChange     =   315
      Left            =   10800
      Max             =   100
      SmallChange     =   315
      TabIndex        =   59
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   58
      Top             =   9615
      Width           =   10815
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   10575
      Begin VB.CheckBox chkExclRET 
         Caption         =   "Exclude Employment Status of RET"
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
         Left            =   120
         TabIndex        =   28
         Tag             =   "Check to Exclude Employees with RET Employment Status"
         Top             =   7380
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.CheckBox chkExclCONP 
         Caption         =   "Exclude Employment Status of CONP"
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
         Left            =   120
         TabIndex        =   27
         Tag             =   "Check to Exclude Employees with CONP Employment Status"
         Top             =   7080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame frTotHrsPos 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3360
         TabIndex        =   62
         Top             =   6285
         Visible         =   0   'False
         Width           =   4155
         Begin Threed.SSOption optCurrHis 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Tag             =   "Type of Report"
            Top             =   0
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "   Current Position only"
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
         Begin Threed.SSOption optCurrHis 
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   64
            TabStop         =   0   'False
            Tag             =   "Type of Report"
            Top             =   0
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "   All Positions"
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
      Begin VB.CheckBox chkFromHireDate 
         Caption         =   "Include hours between Original Hire Date and the 'As of Date'"
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
         Left            =   4680
         TabIndex        =   12
         Top             =   2895
         Visible         =   0   'False
         Width           =   4875
      End
      Begin VB.TextBox txtMonthsTo 
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
         Left            =   3420
         MaxLength       =   7
         TabIndex        =   9
         Tag             =   "10-Enter Seniority Months To"
         Top             =   2600
         Width           =   870
      End
      Begin VB.TextBox txtMonths 
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
         Left            =   1995
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "10-Enter Seniority Months From"
         Top             =   2600
         Width           =   870
      End
      Begin VB.ComboBox cmbSeniority 
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
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Select Seniority Based On"
         Top             =   2220
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
         TabIndex        =   29
         Tag             =   "First Level of grouping records"
         Top             =   8385
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "Second level of grouping records"
         Top             =   8715
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "Final sorting of records - no totals"
         Top             =   9030
         Width           =   2325
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
         Left            =   4725
         TabIndex        =   7
         Top             =   2190
         Width           =   2235
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   20
         Tag             =   "00-Employee Position Shift"
         Top             =   4910
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox cmbAnnMonth 
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
         Left            =   6165
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Select Seniority Based On"
         Top             =   2585
         Visible         =   0   'False
         Width           =   765
      End
      Begin Threed.SSCheck chkIncludeTermFile 
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Tag             =   "Check to included Terminated Employees"
         Top             =   6600
         Visible         =   0   'False
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Include Terminated Employee"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "40-As of Date"
         Top             =   2940
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkShowDet 
         Height          =   225
         Index           =   3
         Left            =   4515
         TabIndex        =   26
         Tag             =   "Hide all detail records - show summaries only"
         Top             =   8955
         Visible         =   0   'False
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "   Show Detail"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         Font3D          =   3
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
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
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
         Index           =   4
         Left            =   1680
         TabIndex        =   14
         Tag             =   "00-Enter Administered By Code"
         Top             =   3600
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   15
         Tag             =   "00-Enter Section Code"
         Top             =   3930
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   13
         Tag             =   "00-Enter Region Code"
         Top             =   3270
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Tag             =   "10-Enter Employee Number"
         Top             =   4260
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkGender 
         Height          =   225
         Index           =   1
         Left            =   3420
         TabIndex        =   19
         Tag             =   "Show Female"
         Top             =   4920
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Female"
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
         Font3D          =   3
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   21
         Tag             =   "ADRE-Attendance Reason"
         Top             =   5200
         Visible         =   0   'False
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkGender 
         Height          =   225
         Index           =   0
         Left            =   1980
         TabIndex        =   18
         Tag             =   "Show Male"
         Top             =   4920
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Male"
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
         Font3D          =   3
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_MACHINE_NUM"
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   23
         Tag             =   "00-Machine #"
         Top             =   5860
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   12
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_PROJECT_CODE"
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   22
         Tag             =   "00-Account Code"
         Top             =   5540
         Visible         =   0   'False
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   11
         LookupType      =   11
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   4590
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
      Begin Threed.SSCheck chkPosTotalHrs 
         Height          =   345
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Visible         =   0   'False
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Total Hours broken by Position"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSenMonthTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   3050
         TabIndex        =   61
         Top             =   2645
         Width           =   195
      End
      Begin VB.Label lblAnnMonth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Anniversary Month"
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
         TabIndex        =   60
         Top             =   2645
         Visible         =   0   'False
         Width           =   1440
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
         TabIndex        =   57
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
         Left            =   120
         TabIndex        =   56
         Top             =   615
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
         TabIndex        =   55
         Top             =   1275
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
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
         TabIndex        =   54
         Top             =   1605
         Width           =   1350
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seniority Based On"
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
         Index           =   5
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   1350
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "As of Date"
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
         Top             =   2985
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Another Version of this FRM in IHR01"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5055
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Label lblSelCri 
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
         TabIndex        =   50
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblRepGrp 
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
         Top             =   8070
         Width           =   1575
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
         Index           =   6
         Left            =   120
         TabIndex        =   48
         Top             =   8430
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
         Left            =   120
         TabIndex        =   47
         Top             =   8745
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
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   9060
         Width           =   825
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
         TabIndex        =   45
         Top             =   945
         Width           =   615
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
         TabIndex        =   44
         Top             =   3315
         Width           =   510
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
         TabIndex        =   43
         Top             =   3645
         Width           =   1125
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
         TabIndex        =   42
         Top             =   3975
         Width           =   540
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
         TabIndex        =   41
         Top             =   4305
         Width           =   1290
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
         TabIndex        =   40
         Top             =   1935
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
         Left            =   5280
         TabIndex        =   39
         Top             =   4955
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblMonths 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seniority Months From"
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
         Top             =   2640
         Width           =   1560
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
         Left            =   120
         TabIndex        =   37
         Top             =   5245
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblGender 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Top             =   4935
         Width           =   645
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         Index           =   23
         Left            =   120
         TabIndex        =   35
         Top             =   5555
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblMachine 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine #"
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
         Left            =   120
         TabIndex        =   34
         Top             =   5875
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
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
         Top             =   4635
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10320
      Top             =   8280
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
Attribute VB_Name = "frmRSeniority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lblSenTitle 'lblSeniority
Dim WithEvents CN001 As ADODB.Connection
Attribute CN001.VB_VarHelpID = -1
Dim WithEvents CN001W As ADODB.Connection
Attribute CN001W.VB_VarHelpID = -1

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkIncludeTermFile_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkPosTotalHrs_Click(Value As Integer)
    If chkPosTotalHrs Then
        frTotHrsPos.Visible = True
    Else
        frTotHrsPos.Visible = False
    End If
End Sub

Private Sub cmbAnnMonth_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbSeniority_Click()
'Seniority by Hours
chkInclAtt.Visible = cmbSeniority.ListIndex = 4
lblAttCodes.Visible = cmbSeniority.ListIndex = 4
clpCode(6).Visible = cmbSeniority.ListIndex = 4
clpCode(7).Visible = cmbSeniority.ListIndex = 4
clpCode(8).Visible = cmbSeniority.ListIndex = 4
lblTitle(23).Visible = cmbSeniority.ListIndex = 4
lblMachine.Visible = cmbSeniority.ListIndex = 4
chkPosTotalHrs.Visible = cmbSeniority.ListIndex = 4
optCurrHis(0).Visible = cmbSeniority.ListIndex = 4
optCurrHis(1).Visible = cmbSeniority.ListIndex = 4

'Ticket #21501
chkFromHireDate.Visible = cmbSeniority.ListIndex = 4

If glbCompSerial = "S/N - 2352W" Then 'Tobias House
    Call DispAnnMonth
End If
If glbCompSerial = "S/N - 2214W" Then    'Casey House - Ticket #15276
    If cmbSeniority.ListIndex = 4 Then
        chkInclAtt.Value = 1
    End If
End If
End Sub

Private Sub cmbSeniority_GotFocus()
  Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Seniority Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
Resume Next

End Sub

Public Sub cmdView_Click()
Dim x%
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
'MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "View Click", "HREMP", "Seniority")
Resume Next

End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim x%
cmbSeniority.AddItem lStr("Original Hire Date")
cmbSeniority.AddItem lStr("Seniority Date")
cmbSeniority.AddItem lStr("Last Hire Date")
cmbSeniority.AddItem lStr("Union Date")
cmbSeniority.AddItem lStr("Seniority Hours")
cmbSeniority.AddItem lStr("First Day")
cmbSeniority.AddItem lStr("User Defined Date") 'added by RAUBREY 5/9/97
'Ticket #22682 - Release 8.0
If Not glbMulti Then
    cmbSeniority.AddItem lStr("Current Position")
End If
cmbSeniority.ListIndex = 0

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem "Position Code"
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem "(none)"

comGroup(1).AddItem "Years of Experience"  'Added 1/16/97
comGroup(1).AddItem lStr("Union")
comGroup(1).AddItem "Employee Name"
comGroup(1).AddItem "(none)"

comGroup(2).AddItem "Employee Number"
comGroup(2).AddItem "Employee Name"

If glbLinamar Then
    comGroup(0).ListIndex = 9
Else
    comGroup(0).ListIndex = 4
End If
comGroup(1).ListIndex = 0
comGroup(2).ListIndex = 0

comGroup(1).Enabled = True  'False    'Ticket #18668
comGroup(2).Enabled = True  'False

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then
        If glbCompSerial = "S/N - 2259W" And comGroup(0) = lStr("Union") And cmbSeniority = lStr("Original Hire Date") Then
            strCd$ = "@txtUnion"
        Else
            strCd$ = IIf(glbMulti, "HR_JOB_HISTORY.JH_ORG", "HREMP.ED_ORG")
        End If
    End If
    If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 3 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_SECTION"
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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

'Hemu 11/01/2004 Begin
If glbCompSerial = "S/N - 2352W" Then   'Tobias House
    EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpPT.Text) & "']"
Else
    EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
End If
'Hemu 11/01/2004 End


If glbCompSerial = "S/N - 2341W" Then
    Select Case cmbSeniority
        Case lStr("Seniority Date")
            EECri = "{HREMP.ED_PT}= 'FT'"
        Case lStr("Seniority Hours")
        EECri = "{HREMP.ED_PT}= 'PT'"
    End Select
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Gender()
Dim gender

If chkGender(0).Value = True And chkGender(1).Value = True Then
'    If glbOracle Then
'        gender = "({HREMP.ED_SEX} IN ['M','F'])"
'    Else
'        gender = "({HREMP.ED_SEX} IN ('M','F'))"
'    End If
    gender = ""
ElseIf chkGender(0).Value = True Then
    gender = "({HREMP.ED_SEX} = 'M')"
ElseIf chkGender(1).Value = True Then
    gender = "({HREMP.ED_SEX} = 'F')"
Else
    gender = "({HREMP.ED_SEX} <> 'F') AND ({HREMP.ED_SEX} <> 'M')"
End If

If Len(gender) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = gender
    Else
        glbstrSelCri = glbstrSelCri & " AND " & gender
    End If
    glbiOneWhere = True
End If
        
End Sub

Private Sub Cri_AnnMonth()
Dim EECri As String, OneSet%, x%
Dim I

If Not cmbAnnMonth.Visible Then Exit Sub
If Len(cmbAnnMonth) = 0 Then Exit Sub

I = cmbAnnMonth.ListIndex

EECri = "month({HREMP.ED_DOH}) = " & I & " and {@Seniority} <=3"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$, xNoFiles, dscGroup$
Dim dtYYY%, dtMM%, dtDD%
Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri

If cmbSeniority = lStr("Seniority Hours") Then
    Call Cri_WRKFile 'laura nov 19, 1997
    glbstrSelCri = "{HRSENHRS.WRKEMP}='" & glbUserID & "'"
Else
    Call glbCri_DeptUN(clpDept.Text)
    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
    For x% = 0 To 5
        Call Cri_Code(x%)
    Next x%
    Call Cri_PT
    Call Cri_Shift
    Call Cri_Job 'Ticket #13238 Frank Jun 21, 2007
    Call Cri_EE
    'Call Cri_Status
    
    'If glbCompSerial = "S/N - 2352W" Then    'Jerry said to open this for everyone Ticket #8920
      Call Cri_Gender
    'End If
    
    If glbCompSerial = "S/N - 2352W" Then 'Tobias House
        Call Cri_AnnMonth
    End If
End If

dscGroup$ = "ASOF = 'as of " & Format(dlpAsOf.Text, "mmmm dd, yyyy") & "'"
Me.vbxCrystal.Formulas(0) = dscGroup$
dtYYY% = Year(dlpAsOf.Text)
dtMM% = month(dlpAsOf.Text)
dtDD% = Day(dlpAsOf.Text)
dscGroup$ = "ASOFDATE = DATE(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"
Me.vbxCrystal.Formulas(1) = dscGroup$

'Hemu - 07/11/2003 Begin - For the Seniority Months >=
If Len(txtMonths.Text) > 0 Then
    If Len(txtMonthsTo.Text) > 0 Then
        Me.vbxCrystal.Formulas(27) = "SenMonths = '" & txtMonths.Text & "'"
        Me.vbxCrystal.Formulas(28) = "SenMonthsTo = '" & txtMonthsTo.Text & "'"
    Else
        Me.vbxCrystal.Formulas(27) = "SenMonths = '" & txtMonths.Text & "'"
    End If
Else
    If Len(txtMonthsTo.Text) > 0 Then
        Me.vbxCrystal.Formulas(27) = "SenMonths = '0'"
        Me.vbxCrystal.Formulas(28) = "SenMonthsTo = '" & txtMonthsTo.Text & "'"
    Else
        Me.vbxCrystal.Formulas(27) = "SenMonths = '0'"
    End If
End If
'Hemu - 07/11/2003 End

xNoFiles = 7

' report name
lblSenTitle = cmbSeniority
Select Case comGroup(0)
Case "(none)"
    Select Case cmbSeniority
    Case lStr("Original Hire Date")
        If glbLinamar Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RLsendo2.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sendo2.rpt"
        End If
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by Original Hire Date")
        GoTo Tag001
    Case lStr("Seniority Date")
        If glbCompSerial = "S/N - 2178W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2178-3.rpt"      'laura febr 11, 1998
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sense2.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Date")
        GoTo Tag001
    Case lStr("Last Hire Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senls2.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by Last Hire Date")
        GoTo Tag001
    Case lStr("Union Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senun2.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by Union Date"
        GoTo Tag001
    Case lStr("Seniority Hours")
        If glbCompSerial = "S/N - 2341W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2341_3.rpt"   'laura febr 11, 1998
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr2.rpt"
            If chkPosTotalHrs Then
                If optCurrHis(0) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr2P.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr2PH.rpt"
                End If
            End If
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Hours")
        GoTo Tag001
    Case lStr("First Day") ' Added by Bryan 07/13/05 Ticket #8874
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senlhn.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by ") & lStr("First Day")
        GoTo Tag001
    Case lStr("User Defined Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senus2.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by User Defined Date"
        GoTo Tag001
    Case lStr("Current Position")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenCPos2.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by Current Position Date"
        GoTo Tag001
    End Select
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senemp.rpt"
    Me.vbxCrystal.WindowTitle = "Seniority Report by Employee Number"
    lblSenTitle = "Employee Number"
Case lStr("Union")
    Select Case cmbSeniority
    Case lStr("Original Hire Date")
        If glbLinamar Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RLsendo1.rpt"
        ElseIf glbCompSerial = "S/N - 2259W" Then 'Oxford
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2259sen.rpt" 'Bryan 06/Apr/06
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sendo1.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Original Hire Date")
        GoTo Tag001
    Case lStr("Seniority Date")
        If glbCompSerial = "S/N - 2341W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2341_2.rpt"   'laura febr 11, 1998
        ElseIf glbCompSerial = "S/N - 2178W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2178-2.rpt"   'laura febr 11, 1998
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sense1.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Date")
        GoTo Tag001
    Case lStr("Last Hire Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senls1.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by Last Hire Date")
        GoTo Tag001
    Case lStr("Union Date")
        'If glbCompSerial = "S/N - 2190W" Then   'Ticket #17569 - WDDS
        '    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2190_rzsenun1.rpt"
        'Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senun1.rpt"
        'End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Union Date")
        GoTo Tag001
    Case lStr("Seniority Hours")
        If glbCompSerial = "S/N - 2341W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2341_3.rpt"   'laura febr 11, 1998
        'ElseIf glbCompSerial = "S/N - 2190W" Then   'Ticket #17569 - WDDS
        '    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2190_rzsenhr1.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr1.rpt"
            If chkPosTotalHrs Then
                If optCurrHis(0) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr1P.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhr1PH.rpt"
                End If
            End If
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Hours")
        GoTo Tag001
    Case lStr("First Day") ' Added by Bryan 07/13/05 Ticket #8874
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senlhu.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by ") & lStr("First Day")
        GoTo Tag001
    Case lStr("User Defined Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senus1.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by User Defined Date"
        GoTo Tag001
    Case lStr("Current Position")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenCPos1.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by Current Position Date"
        GoTo Tag001
    End Select
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senemp.rpt"
    Me.vbxCrystal.WindowTitle = "Seniority Report by Employee Number"
    lblSenTitle = "Employee Number"
Case Else
    Select Case cmbSeniority
    Case lStr("Original Hire Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sendoh.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by Original Hire Date")
        GoTo Tag001
    Case lStr("Seniority Date")
        If glbCompSerial = "S/N - 2341W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2341_2.rpt"   'laura febr 11, 1998
        ElseIf glbCompSerial = "S/N - 2178W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2178-1.rpt"    'laura febr 11, 1998
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "sensen.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Date")
        GoTo Tag001
    Case lStr("Last Hire Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senlst.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by Last Hire Date")
        GoTo Tag001
    Case lStr("Union Date")
        'If glbCompSerial = "S/N - 2190W" Then   'Ticket #17569 - WDDS
        '    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2190_rzsenuni.rpt"
        'Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senuni.rpt"
        'End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Union Date")
        GoTo Tag001
    Case lStr("Seniority Hours")
        If glbCompSerial = "S/N - 2341W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2341_3.rpt"   'laura febr 11, 1998
        'ElseIf glbCompSerial = "S/N - 2190W" Then   'Ticket #17569 - WDDS
        '    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2190_rzsenhrs.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhrs.rpt"
            If chkPosTotalHrs Then
                If optCurrHis(0) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhrsP.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senhrsPH.rpt"
                End If
            End If
        End If
        Me.vbxCrystal.WindowTitle = "Seniority Report by " & lStr("Seniority Hours")
        GoTo Tag001
    Case lStr("First Day") ' Added by Bryan 07/13/05 Ticket #8874
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senlh.rpt"
        Me.vbxCrystal.WindowTitle = lStr("Seniority Report by ") & lStr("First Day")
        GoTo Tag001
    Case lStr("User Defined Date")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senusr.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by User Defined Date"
        GoTo Tag001
    Case lStr("Current Position")
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenCPos.rpt"
        Me.vbxCrystal.WindowTitle = "Seniority Report by Current Position Date"
        GoTo Tag001
    End Select
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "senemp.rpt"
    Me.vbxCrystal.WindowTitle = "Seniority Report by Employee Number"
    lblSenTitle = "Employee Number"
End Select
Tag001:

' set to sorting/grouping criteria
If Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenhrs.rpt" Then
    xNoFiles = 11
ElseIf Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenhr1.rpt" Or Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsenhr2.rpt" Then
    xNoFiles = 10
Else
    xNoFiles = xNoFiles + 1
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

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

x% = Cri_Sorts()   'Laura nov 19, 1997

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    If cmbSeniority = lStr("Seniority Hours") Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To xNoFiles - 1
            If x% = 6 Then
                Me.vbxCrystal.DataFiles(6) = glbIHRDBW
            Else
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next x%
    Else
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
End If

Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "seniority", "Seniority Report", "Select")
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
Dim senHours As Boolean

Cri_Sorts = 0

Call setRptLabel(Me, 0)

Select Case cmbSeniority
Case lStr("Original Hire Date")
    Me.vbxCrystal.Formulas(18) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
Case lStr("Seniority Date")
    Me.vbxCrystal.Formulas(19) = "lblSeniority='" & lStr("Seniority Date") & "'"
Case lStr("Last Hire Date")
    Me.vbxCrystal.Formulas(20) = "lblLHireDate='" & lStr("Last Hire Date") & "'"
Case lStr("Union Date")
    Me.vbxCrystal.Formulas(21) = "lblUnionDate='" & lStr("Union Date") & "'"
Case lStr("First Day")
    Me.vbxCrystal.Formulas(18) = "lblFirstDay='" & lStr("First Day") & "'"
Case lStr("Seniority Hours")
    Me.vbxCrystal.Formulas(18) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
    senHours = True
    If comGroup(0).Text = lStr("Union") Then Exit Function
    If comGroup(0).Text = "(none)" Then Exit Function
Case lStr("User Defined Date")
    Me.vbxCrystal.Formulas(25) = "lblUserDate='" & lStr("User Defined Date") & "'"
'Case lStr("Current Position")
End Select
Me.vbxCrystal.Formulas(26) = "lblTitle='Seniority Report by " & lblSenTitle & "'"

grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0) = lStr("Union") Then
    If glbCompSerial = "S/N - 2259W" And cmbSeniority = lStr("Original Hire Date") Then
        grpField$ = "{@txtUnion}"
    ElseIf glbMulti Then
        grpField$ = "{HR_JOB_HISTORY.JH_ORG}"
    Else
        grpField$ = "{HREMP.ED_ORG}"
    End If
End If
If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(2) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    Me.vbxCrystal.SectionFormat(z%) = "GH1;T;X;X;X;X;X;X"
    z% = z% + 1
    Me.vbxCrystal.SectionFormat(z%) = "GF1;T;X;X;X;X;X;X"
    z% = z% + 1
    Me.vbxCrystal.SectionFormat(z%) = "GF2;F;X;X;X;X;X;X"
    z% = z% + 1
    
    If senHours Then Exit Function
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@Seniority}"
        Case 1: grpField$ = "{HREMP.ED_ORG}"
        Case 2: grpField$ = "{@EFullName}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        If GrpIdx% = 0 Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;D"
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREMP.ED_EMPNBR}"
            Case 1: grpField$ = "{@EFullName}"
            Case Else: grpField$ = "(none)"
        End Select
        grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        If grpField$ <> "(none)" Then
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    Else
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREMP.ED_EMPNBR}"
            Case 1: grpField$ = "{@EFullName}"
            Case Else: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        End If
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@Seniority}"
        Case 1: grpField$ = "{HREMP.ED_ORG}"
        Case 2: grpField$ = "{@EFullName}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        If GrpIdx% = 0 Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;D"
        Else
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREMP.ED_EMPNBR}"
            Case 1: grpField$ = "{@EFullName}"
            Case Else: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        End If
    Else
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREMP.ED_EMPNBR}"
            Case 1: grpField$ = "{@EFullName}"
            Case Else: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        End If
    End If
End If

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
Dim EECri As String

If Len(clpCode(2).Text) > 0 Then
    EECri = "{HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "'] "
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

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known"), , "info:HR"
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known", , "info:HR"
    'clpDept.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If clpCode(6).Visible Then
    If Not clpCode(6).ListChecker Then Exit Function
End If

If glbCompSerial = "S/N - 2352W" Then
    If Not clpPT.ListChecker Then
        Exit Function
    End If
Else
    If Not clpPT.ListChecker Then
    'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
        'MsgBox lStr("Category code must be valid")
        'clpPT.SetFocus
        Exit Function
    End If
End If

'Hemu - 07/10/2003 Begin - For Seniority Months >=
If Len(txtMonths.Text) > 0 Then
    If Not IsNumeric(txtMonths.Text) Then
        MsgBox lStr("If Seniority Months From is entered it must be numeric"), , "info:HR"
        txtMonths.SetFocus
        Exit Function
    End If
End If
If Len(txtMonthsTo.Text) > 0 Then
    If Not IsNumeric(txtMonthsTo.Text) Then
        MsgBox lStr("If Seniority Months To is entered it must be numeric"), , "info:HR"
        txtMonthsTo.SetFocus
        Exit Function
    End If
End If
'Hemu - 07/10/2003 End


If Len(dlpAsOf.Text) > 0 Then
    If Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid As of Date", , "info:HR"
        dlpAsOf.SetFocus
        Exit Function
    End If
Else
  dlpAsOf.Text = Format(Now, "Short Date")
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

Me.Caption = lStr(Me.Caption)

Call setCaption(lblTitle(23))
Call setCaption(lblMachine)

lblAttCodes.Caption = lStr("Reason")
chkFromHireDate.Caption = "Include hours between '" & lStr("Original Hire Date") & "' and 'As of Date'"

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If glbCompSerial = "S/N - 2352W" Then
    clpPT.MultiSelect = True
    clpPT.ShowDescription = False
    clpPT.MaxLength = 0
    clpPT.Width = 7515
    
    clpPT.ShowDescription = True
    clpPT.ShowDescription = False
    
    lblGender.Visible = True
    chkGender(0).Visible = True
    chkGender(1).Visible = True
End If

'Ticket# 9981, see note in comAnnMonthAdding
If glbCompSerial = "S/N - 2352W" Then 'Tobias House
    lblAnnMonth.Visible = True
    cmbAnnMonth.Visible = True
    Call comAnnMonthAdding
End If

'Mitchell Plastics - Ticket #22614
If glbCompSerial = "S/N - 2335W" Then
    cmbSeniority.ListIndex = 1
End If

If glbMulti Then lblUnion.ForeColor = &HC000C0

Call INI_Controls(Me)

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob.TransDiv = glbWFCUserSecList
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_Resize()
scrFrame.Height = 9255
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 9000 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7600 Then
            scrControl.Max = 5500
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
    If Me.Width >= 9500 Then
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

Private Sub Cri_WRKFile()
Dim SQLQ, SQLX, WSQLQ, SQLH, SQLTH
Dim rsTT As New ADODB.Recordset
Dim xFld As String

On Error GoTo Sen_WrkError

'Hemu
Set CN001 = New ADODB.Connection
CN001.CommandTimeout = 600
CN001.Open glbAdoIHRDB

Set CN001W = New ADODB.Connection
CN001W.CommandTimeout = 600
CN001W.Open glbAdoIHRDBW

CN001.BeginTrans
CN001.Execute "DELETE FROM HRSENHRS " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
CN001.CommitTrans

'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute "DELETE FROM HRSENHRS " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
'gdbAdoIhr001.CommitTrans
'Hemu

WSQLQ = glbSeleDeptUn
If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION='" & clpCode(5).Text & "'"
If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

If chkGender(0).Value = True And chkGender(1).Value = True Then
    'do nothing - retrieves all by default
ElseIf chkGender(0).Value = True Then
    WSQLQ = WSQLQ & " AND HREMP.ED_SEX = 'M'"
ElseIf chkGender(1).Value = True Then
    WSQLQ = WSQLQ & " AND HREMP.ED_SEX = 'F'"
Else
    WSQLQ = WSQLQ & " AND HREMP.ED_SEX <> 'M' AND HREMP.ED_SEX <> 'F'"
End If

If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

'Ticket #13238 Frank Jun 21, 2007
If Len(clpJob.Text) > 0 Then
    'Ticket #23973 - Added Position History option
    If chkPosTotalHrs And optCurrHis(1) Then
        xFld = "(HR_ATTENDANCE.AD_JOB in  ('" & Replace(clpJob.Text, ",", "','") & "'))"
        WSQLQ = WSQLQ & " AND ED_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE " & xFld & ") "
    Else
        xFld = "(HR_JOB_HISTORY.JH_JOB in  ('" & Replace(clpJob.Text, ",", "','") & "'))"
        WSQLQ = WSQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT <>0 AND " & xFld & ") "
    End If
End If

SQLQ = "SELECT '001' AS COMPANY,ED_EMPNBR,'EDOR' AS UNION_TBL,'EDEM' AS EMPLSTAT_TBL,"
If glbOracle Then
    SQLQ = SQLQ & "SUBSTR(ED_SURNAME || ', ' || ED_FNAME,1,39) AS NAME,"
Else
    SQLQ = SQLQ & "LEFT(ED_SURNAME +', '+ED_FNAME,39) AS NAME,"
End If
SQLQ = SQLQ & "ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH,0 AS NUMBHRS "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
If chkPosTotalHrs Then
    'Ticket #23973 - Added Position History option
    If optCurrHis(0) Then
        SQLQ = SQLQ & ",JH_JOB "
    Else
        SQLQ = SQLQ & ",AD_JOB "
    End If
End If
SQLQ = SQLQ & "FROM HREMP"

If chkPosTotalHrs Then
    'Ticket #23973 - Added Position History option
    If optCurrHis(0) Then
        SQLQ = SQLQ & ", HR_JOB_HISTORY"
    Else
        SQLQ = SQLQ & ", HR_ATTENDANCE"
    End If
End If

SQLQ = SQLQ & in_SQL(glbIHRDB)

SQLX = "INSERT INTO HRSENHRS "
'SQLX = SQLX & in_SQL(glbIHRDB)
SQLX = SQLX & "(COMPANY,EMPLNUMB,UNION_TBL,EMPLSTAT_TBL,NAME,"
SQLX = SQLX & "DIVISION,DEPTNO,UNIONS,EMPLSTAT,DOH,NUMBHRS,WRKEMP"
If chkPosTotalHrs Then
    SQLX = SQLX & ",JOB"
End If
SQLX = SQLX & ") "
SQLX = SQLX & SQLQ & " WHERE " & WSQLQ

If chkPosTotalHrs Then
    'Ticket #23973 - Added Position History option
    If optCurrHis(0) Then
        SQLX = SQLX & " AND ED_EMPNBR = JH_EMPNBR AND JH_CURRENT <> 0"
        SQLX = SQLX & " group by ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH, JH_JOB "
        
        SQLTH = SQLTH & " AND ED_EMPNBR = JH_EMPNBR AND JH_CURRENT <> 0"
        SQLTH = SQLTH & " group by ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH, JH_JOB "
    Else
        SQLX = SQLX & " AND ED_EMPNBR = AD_EMPNBR"
        SQLTH = SQLTH & " AND ED_EMPNBR = AD_EMPNBR"
        
        If Len(clpJob.Text) > 0 Then
            SQLX = SQLX & " AND (HR_ATTENDANCE.AD_JOB in  ('" & Replace(clpJob.Text, ",", "','") & "'))"
            SQLTH = SQLTH & " AND (HR_ATTENDANCE.AD_JOB in  ('" & Replace(clpJob.Text, ",", "','") & "'))"
        End If
        SQLX = SQLX & " group by ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH, AD_JOB "
        SQLTH = SQLTH & " group by ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH, AD_JOB "
    End If
End If

'Hemu
'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute SQLQ
'gdbAdoIhr001.CommitTrans

'Ticket #22682 - Seniority by Hours was not showing the Totals by Position correctly when Include Attend. History
'was selected. This will add Positions that are associated to the Attendance records in the history.
If chkInclAtt Then
    SQLH = SQLQ
    SQLH = SQLH & " WHERE " & WSQLQ & SQLTH
    SQLH = Replace(Replace(SQLH, "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY"), "AD_", "AH_")
    
    CN001W.BeginTrans
    CN001W.Execute SQLX & " UNION " & SQLH
    CN001W.CommitTrans
Else
    CN001W.BeginTrans
    CN001W.Execute SQLX
    CN001W.CommitTrans
End If
'Hemu


If glbOracle Then
    SQLQ = "SELECT SUM(AD_HRS) FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_SEN<>0 "
    
    'Ticket #21501
    If chkFromHireDate Then
        SQLQ = SQLQ & " AND (AD_DOA>=(SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = AD_EMPNBR)"
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpAsOf.Text) & ")"
    Else
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpAsOf.Text)
    End If
    
    If clpCode(6).Text <> "" Then
        SQLQ = SQLQ & " AND AD_REASON IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
    End If
    If Len(clpCode(7).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_MACHINE_NUM='" & clpCode(7).Text & "'"
    End If
    If Len(clpCode(8).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_PROJECT_CODE='" & clpCode(8).Text & "'"
    End If
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & WSQLQ & ")"

    SQLX = " Update HRSENHRS SET NUMBHRS =(" & SQLQ & " AND HR_ATTENDANCE.AD_EMPNBR=HRSENHRS.EMPLNUMB)"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLX
    gdbAdoIhr001.CommitTrans
Else
    SQLQ = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS NUM1"
    'Ticket #18000
    If chkPosTotalHrs Then
        SQLQ = SQLQ & ", AD_JOB"
    End If
    SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    
    SQLQ = SQLQ & " WHERE AD_SEN<>0 "
    
    'Ticket #21501
    If chkFromHireDate Then
        SQLQ = SQLQ & " AND (AD_DOA>=(SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = AD_EMPNBR)"
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpAsOf.Text) & ")"
    Else
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpAsOf.Text)
    End If
    
    If clpCode(6).Text <> "" Then
        SQLQ = SQLQ & " AND AD_REASON IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
    End If
    If Len(clpCode(7).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_MACHINE_NUM='" & clpCode(7).Text & "'"
    End If
    If Len(clpCode(8).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_PROJECT_CODE='" & clpCode(8).Text & "'"
    End If
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " WHERE " & WSQLQ & ")"
    
    'Ticket #18000
    If chkPosTotalHrs And glbLambton Then
        'Ticket #23973 - Added Position History option
        If optCurrHis(0) Then
            SQLQ = SQLQ & " AND ad_job in (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1)  from hr_job_history where jh_current <> 0 and jh_empnbr = ad_empnbr)"
        Else
            SQLQ = SQLQ & " AND ad_job in (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1)  from hr_job_history where jh_empnbr = ad_empnbr)"
        End If
    ElseIf chkPosTotalHrs Then
        'Ticket #23973 - Added Position History option
        If optCurrHis(0) Then
            SQLQ = SQLQ & " AND ad_job in (select jh_job from hr_job_history " & in_SQL(glbIHRDB) & " where jh_current <> 0 and jh_empnbr = ad_empnbr)"
        End If
    End If
        
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    
    'Ticket #18000
    If chkPosTotalHrs Then
        SQLQ = SQLQ & ", AD_JOB"
    End If
    
    'Update with total seniority hours
    If glbSQL Then
        SQLX = " Update HRSENHRS "
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1"
        
        'Ticket #18000
        'If chkPosTotalHrs Then
        '    SQLX = SQLX & ", JOB = TRSNUM.AD_JOB"
        'End If
        
        SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        
        'Ticket #18000
        If chkPosTotalHrs And glbLambton Then
            'Ticket #23973 - Added Position History option
            If optCurrHis(0) Then
                SQLX = SQLX & " and TRSNUM.AD_JOB IN (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1) from hr_job_history where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
            Else
                SQLX = SQLX & " and TRSNUM.AD_JOB IN (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1) from hr_job_history where jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
            End If
        ElseIf chkPosTotalHrs Then
            'Ticket #23973 - Added Position History option
            If optCurrHis(0) Then
                SQLX = SQLX & " and TRSNUM.AD_JOB IN (select jh_job from hr_job_history where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
            Else
                SQLX = SQLX & " and (TRSNUM.AD_JOB = HRSENHRS.JOB OR (TRSNUM.AD_JOB IS NULL AND HRSENHRS.JOB IS NULL))"
            End If
        End If
        
        gdbAdoIhr001.Execute SQLX
    Else
        'Hemu
        'gdbAdoIhr001.Execute "DELETE FROM TRSNUM " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
        CN001W.BeginTrans
        CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
        CN001W.CommitTrans
        'Hemu
        
        'Ticket #18000
        If chkPosTotalHrs Then
            SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,AD_JOB,WRKEMP) " & SQLQ
        Else
            SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
        End If
        'Hemu
        'gdbAdoIhr001.Execute SQLX
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
                
        SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1 "
        SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
        SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
        
        'Ticket #18000
        If chkPosTotalHrs Then
            'Ticket #23973 - Added Position History option
            If optCurrHis(0) Then
                SQLX = SQLX & " and TRSNUM.AD_JOB IN (select jh_job from hr_job_history " & in_SQL(glbIHRDB) & " where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
            Else
                SQLX = SQLX & " and (TRSNUM.AD_JOB = HRSENHRS.JOB OR (TRSNUM.AD_JOB IS NULL AND HRSENHRS.JOB IS NULL))"
            End If
        End If
        
        'Hemu
        'gdbAdoIhr001W.Execute SQLX
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
    End If
End If

If chkInclAtt Then
    If glbOracle Then
        SQLQ = "SELECT SUM(AH_HRS) FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_SEN<>0 "
        
        'Ticket #21501
        If chkFromHireDate Then
            SQLQ = SQLQ & " AND (AH_DOA>=(SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = AH_EMPNBR)"
            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(dlpAsOf.Text) & ")"
        Else
            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(dlpAsOf.Text)
        End If
        
        If clpCode(6).Text <> "" Then
            SQLQ = SQLQ & " AND AH_REASON IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
        End If
        If Len(clpCode(7).Text) > 0 Then
            SQLQ = SQLQ & " AND AH_MACHINE_NUM='" & clpCode(7).Text & "'"
        End If
        If Len(clpCode(8).Text) > 0 Then
            SQLQ = SQLQ & " AND AH_PROJECT_CODE='" & clpCode(8).Text & "'"
        End If
        SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & WSQLQ & ")"
        
        SQLX = " Update HRSENHRS SET NUMBHRS =NUMBHRS+(" & SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_EMPNBR=HRSENHRS.EMPLNUMB)"
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLX
        gdbAdoIhr001.CommitTrans
    Else
        SQLQ = "SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS NUM1"
        
        'Ticket #18000
        If chkPosTotalHrs Then
            SQLQ = SQLQ & ", AH_JOB"
        End If
        
        SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE AH_SEN<>0 "
        
        'Ticket #21501
        If chkFromHireDate Then
            SQLQ = SQLQ & " AND (AH_DOA>=(SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = AH_EMPNBR)"
            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(dlpAsOf.Text) & ")"
        Else
            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(dlpAsOf.Text)
        End If
        
        If clpCode(6).Text <> "" Then
            SQLQ = SQLQ & " AND AH_REASON IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
        End If
        If Len(clpCode(7).Text) > 0 Then
            SQLQ = SQLQ & " AND AH_MACHINE_NUM='" & clpCode(7).Text & "'"
        End If
        If Len(clpCode(8).Text) > 0 Then
            SQLQ = SQLQ & " AND AH_PROJECT_CODE='" & clpCode(8).Text & "'"
        End If
        SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE " & WSQLQ & ")"

        'Ticket #18000
        If chkPosTotalHrs And glbLambton Then
            'Ticket #23973 - Added Position History option
            If optCurrHis(0) Then
                SQLQ = SQLQ & " AND ah_job in (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1)  from hr_job_history where jh_current <> 0 and jh_empnbr = ah_empnbr)"
            Else
                SQLQ = SQLQ & " AND ah_job in (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1)  from hr_job_history where jh_empnbr = ah_empnbr)"
            End If
        ElseIf chkPosTotalHrs Then
            'Ticket #23973 - Added Position History option
            If optCurrHis(0) Then
                SQLQ = SQLQ & " AND ah_job in (select jh_job from hr_job_history " & in_SQL(glbIHRDB) & " where jh_current <> 0 and jh_empnbr = ah_empnbr)"
            End If
        End If

        SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
       
        'Ticket #18000
        If chkPosTotalHrs Then
            SQLQ = SQLQ & ", AH_JOB"
        End If
       
        
        'Update with total seniority hours
        If glbSQL Then
            SQLX = " Update HRSENHRS "
            SQLX = SQLX & " SET NUMBHRS = ISNULL(NUMBHRS,0)+TRSNUM.NUM1"
            
            'Ticket #18000
            'If chkPosTotalHrs Then
            '    SQLX = SQLX & ", JOB = TRSNUM.AD_JOB"
            'End If
            
            SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            
            'Ticket #18000
            If chkPosTotalHrs And glbLambton Then
                'Ticket #23973 - Added Position History option
                If optCurrHis(0) Then
                    SQLX = SQLX & " and TRSNUM.AH_JOB = (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1) from hr_job_history where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
                Else
                    SQLX = SQLX & " and TRSNUM.AH_JOB = (select substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1) from hr_job_history where jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
                End If
            ElseIf chkPosTotalHrs Then
                'Ticket #23973 - Added Position History option
                If optCurrHis(0) Then
                    SQLX = SQLX & " and TRSNUM.AH_JOB = (select jh_job from hr_job_history where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
                Else
                    SQLX = SQLX & " and (TRSNUM.AH_JOB = HRSENHRS.JOB OR (TRSNUM.AH_JOB IS NULL AND HRSENHRS.JOB IS NULL))"
                End If
            End If
            
            gdbAdoIhr001.Execute SQLX
        Else
            'Hemu
            'gdbAdoIhr001.Execute "DELETE FROM TRSNUM " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
            CN001W.BeginTrans
            CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
            CN001W.CommitTrans
            'Hemu
            
            'SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & in_SQL(glbIHRDBW) & SQLQ
            If chkPosTotalHrs Then
                SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,AD_JOB,WRKEMP) " & SQLQ
            Else
                SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
            End If
            
            'Hemu
            'gdbAdoIhr001.Execute SQLX
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
            
            SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            SQLX = SQLX & " SET NUMBHRS = NUMBHRS + TRSNUM.NUM1 "
            SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
            SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
            
            'Ticket #18000
            If chkPosTotalHrs Then
                'Ticket #23973 - Added Position History option
                If optCurrHis(0) Then
                    SQLX = SQLX & " and TRSNUM.AD_JOB = (select jh_job from hr_job_history " & in_SQL(glbIHRDB) & " where jh_current <> 0 and jh_empnbr = HRSENHRS.EMPLNUMB and jh_job = HRSENHRS.JOB )"
                Else
                    SQLX = SQLX & " and (TRSNUM.AD_JOB = HRSENHRS.JOB OR (TRSNUM.AD_JOB IS NULL AND HRSENHRS.JOB IS NULL))"
                End If
            End If
            
            'Hemu
            'gdbAdoIhr001W.Execute SQLX
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
        End If
    End If
End If
'-------------------
'SQLQ = "update HRSENHRS set job = (select jh_job from hr_job_history where jh_current <> 0 and jh_empnbr = emplnumb and substring(jh_grid,1,1) + jh_job + substring(jh_grid,2,1) = job  )"
'CN001W.BeginTrans
'CN001W.Execute SQLQ
'CN001W.CommitTrans
'-------------------

CN001.Close
CN001W.Close

Set CN001 = Nothing
Set CN001W = Nothing

Exit Sub

Sen_WrkError:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtMonths_GotFocus()
Call SetPanHelp(Me.ActiveControl)
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

Private Sub comAnnMonthAdding()
'Modify the Seniority Report by Original Date of Hire. Add a selection criteria asking for "Anniversay Month".
'When selected by the users, the report will only show employees who have their Original Date of Hire month
'equal the Anniversary Month and have their Years of Experience less than or equal to 3 years.
    cmbAnnMonth.AddItem ""
    cmbAnnMonth.AddItem "Jan"
    cmbAnnMonth.AddItem "Feb"
    cmbAnnMonth.AddItem "Mar"
    cmbAnnMonth.AddItem "Apr"
    cmbAnnMonth.AddItem "May"
    cmbAnnMonth.AddItem "Jun"
    cmbAnnMonth.AddItem "Jul"
    cmbAnnMonth.AddItem "Aug"
    cmbAnnMonth.AddItem "Sep"
    cmbAnnMonth.AddItem "Oct"
    cmbAnnMonth.AddItem "Nov"
    cmbAnnMonth.AddItem "Dec"
End Sub

Private Sub DispAnnMonth()
If cmbSeniority = lStr("Original Hire Date") Then
    lblAnnMonth.Visible = True
    cmbAnnMonth.Visible = True
Else
    lblAnnMonth.Visible = False
    cmbAnnMonth.Visible = False
End If
End Sub

Private Sub Cri_Job()
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim x%

If Len(clpJob.Text) = 0 Then Exit Sub

fld$ = "HR_JOB_HISTORY.JH_JOB"

CodeCri = "({" & fld$ & "} in  ['" & Replace(clpJob.Text, ",", "','") & "'])"

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub
