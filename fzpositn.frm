VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRPosition 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   11475
   ClientLeft      =   435
   ClientTop       =   870
   ClientWidth     =   13515
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11475
   ScaleWidth      =   13515
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   10935
      LargeChange     =   300
      Left            =   13080
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   100
      Top             =   120
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   11295
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Width           =   12975
      _Version        =   65536
      _ExtentX        =   22886
      _ExtentY        =   19923
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
      BevelOuter      =   1
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   11055
         Left            =   0
         ScaleHeight     =   11025
         ScaleWidth      =   12825
         TabIndex        =   75
         Top             =   120
         Width           =   12855
         Begin VB.CheckBox chkExclCONP 
            Caption         =   "Exclude Employment Status of CONP"
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Tag             =   "Check to Exclude Employees with CONP Employment Status"
            Top             =   6600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.CheckBox chkExclRET 
            Caption         =   "Exclude Employment Status of RET"
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Tag             =   "Check to Exclude Employees with RET Employment Status"
            Top             =   6855
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.Frame frRptGrouping 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   111
            Top             =   9600
            Width           =   4455
            Begin VB.ComboBox comGroup 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   0
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Tag             =   "First Level of grouping records"
               Top             =   285
               Width           =   2325
            End
            Begin VB.ComboBox comGroup 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   1
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Tag             =   "Final sorting of records"
               Top             =   600
               Width           =   2325
            End
            Begin VB.ComboBox comGroup 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Tag             =   "Final sorting of records"
               Top             =   900
               Width           =   2325
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
               TabIndex        =   115
               Top             =   0
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
               Left            =   120
               TabIndex        =   114
               Top             =   315
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
               Left            =   120
               TabIndex        =   113
               Top             =   645
               Width           =   660
            End
            Begin VB.Label lblGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Work History Sort"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   112
               Top             =   930
               Width           =   1695
            End
         End
         Begin VB.Frame frHideShowDetails 
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   120
            TabIndex        =   110
            Top             =   7200
            Visible         =   0   'False
            Width           =   12495
            Begin VB.CheckBox chkShowMedical 
               Caption         =   "Show Medical Contacts"
               Height          =   285
               Left            =   0
               TabIndex        =   48
               Top             =   480
               Value           =   1  'Checked
               Width           =   2085
            End
            Begin VB.CheckBox chkShowDependents 
               Caption         =   "Show Dependents"
               Height          =   285
               Left            =   3120
               TabIndex        =   54
               Top             =   240
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowCurPosition 
               Caption         =   "Show Current Position"
               Height          =   285
               Left            =   0
               TabIndex        =   51
               Top             =   1320
               Value           =   1  'Checked
               Width           =   2085
            End
            Begin VB.CheckBox chkShowCurSalary 
               Caption         =   "Show Current Salary"
               Height          =   285
               Left            =   0
               TabIndex        =   52
               Top             =   1560
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowCurPerform 
               Caption         =   "Show Current Performance"
               Height          =   285
               Left            =   0
               TabIndex        =   53
               Top             =   1800
               Value           =   1  'Checked
               Width           =   2205
            End
            Begin VB.CheckBox chkShowPositionHist 
               Caption         =   "Show Position History"
               Height          =   285
               Left            =   3120
               TabIndex        =   58
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowSalaryHist 
               Caption         =   "Show Salary History"
               Height          =   285
               Left            =   3120
               TabIndex        =   59
               Top             =   1560
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowPerformHist 
               Caption         =   "Show Performance History"
               Height          =   285
               Left            =   3120
               TabIndex        =   60
               Top             =   1800
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowLanguages 
               Caption         =   "Show Languages"
               Height          =   285
               Left            =   3120
               TabIndex        =   55
               Top             =   480
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowSkills 
               Caption         =   "Show Skills"
               Height          =   285
               Left            =   3120
               TabIndex        =   56
               Top             =   720
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowOtherEarnings 
               Caption         =   "Show Other Earnings"
               Height          =   285
               Left            =   3120
               TabIndex        =   57
               Top             =   960
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.CheckBox chkShowFormalEdu 
               Caption         =   "Show Formal Education"
               Height          =   285
               Left            =   5880
               TabIndex        =   61
               Top             =   240
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowCourseSeminars 
               Caption         =   "Show Courses/Seminars"
               Height          =   285
               Left            =   5880
               TabIndex        =   62
               Top             =   480
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowAssociations 
               Caption         =   "Show Associations"
               Height          =   285
               Left            =   5880
               TabIndex        =   63
               Top             =   720
               Value           =   1  'Checked
               Width           =   2685
            End
            Begin VB.CheckBox chkShowBenefits 
               Caption         =   "Show Benefits Information"
               Height          =   285
               Left            =   5880
               TabIndex        =   64
               Top             =   960
               Value           =   1  'Checked
               Width           =   2325
            End
            Begin VB.CheckBox chkShowBeneficiary 
               Caption         =   "Show Beneficiary Information"
               Height          =   285
               Left            =   5880
               TabIndex        =   65
               Top             =   1200
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowVacSickCompWSIB 
               Caption         =   "Show Vacation/Sick, Comp. Time && WSIB Hours"
               Height          =   285
               Left            =   5880
               TabIndex        =   66
               Top             =   1440
               Value           =   1  'Checked
               Width           =   3885
            End
            Begin VB.CheckBox chkShowHourlyEntitle 
               Caption         =   "Show Hourly Entitlements in Hours"
               Height          =   285
               Left            =   5880
               TabIndex        =   67
               Top             =   1680
               Value           =   1  'Checked
               Width           =   2925
            End
            Begin VB.CheckBox chkShowDollarEntitle 
               Caption         =   "Show Dollar Entitlements"
               Height          =   285
               Left            =   10080
               TabIndex        =   68
               Top             =   240
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowEmploymentHist 
               Caption         =   "Show Employment History"
               Height          =   285
               Left            =   10080
               TabIndex        =   69
               Top             =   480
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowLeaveHist 
               Caption         =   "Show Leave History"
               Height          =   285
               Left            =   10080
               TabIndex        =   70
               Top             =   720
               Value           =   1  'Checked
               Width           =   2445
            End
            Begin VB.CheckBox chkShowEmergency 
               Caption         =   "Show Emergency Contacts"
               Height          =   285
               Left            =   0
               TabIndex        =   47
               Top             =   240
               Value           =   1  'Checked
               Width           =   2685
            End
            Begin VB.CheckBox chkShowBanking 
               Caption         =   "Show Payroll/Banking Information"
               Height          =   285
               Left            =   0
               TabIndex        =   49
               Top             =   720
               Value           =   1  'Checked
               Width           =   2805
            End
            Begin VB.CheckBox chkShowEmploymentInfo 
               Caption         =   "Show Employment Information"
               Height          =   285
               Left            =   0
               TabIndex        =   50
               Top             =   960
               Value           =   1  'Checked
               Width           =   2685
            End
            Begin VB.CheckBox chkShowPersonalInfo 
               Caption         =   "Show Personal Information"
               Height          =   285
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Value           =   1  'Checked
               Width           =   2685
            End
         End
         Begin VB.CheckBox chkTotalsOnly 
            Caption         =   "Show Totals Only"
            Height          =   285
            Left            =   4560
            TabIndex        =   41
            Top             =   6300
            Width           =   1635
         End
         Begin VB.ComboBox cmbAnnMonth 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2050
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Tag             =   "Select Anniversary Month"
            Top             =   5885
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CheckBox chkIncludeTerm 
            Caption         =   "Include Terminated Employees"
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   6300
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.ComboBox comCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10800
            TabIndex        =   27
            Tag             =   "00-Country"
            Top             =   4875
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.CheckBox chkForAudit 
            Caption         =   "For Data Audit"
            Height          =   285
            Left            =   9240
            TabIndex        =   43
            Top             =   6300
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkWeeklyEmpList 
            Caption         =   "Show Weekly Employee List"
            Height          =   285
            Left            =   6480
            TabIndex        =   42
            Top             =   6300
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.CheckBox chkLastDay 
            Caption         =   "Show Last Day"
            Height          =   285
            Left            =   2880
            TabIndex        =   40
            Top             =   6300
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtShift 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8380
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "00-Employee Position Shift"
            Top             =   4556
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.CheckBox chkVDesc 
            Caption         =   "View Languages Descriptions"
            Height          =   285
            Left            =   6885
            TabIndex        =   32
            Top             =   5220
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.Frame frmDesc 
            Caption         =   "Languages Description"
            Height          =   1485
            Left            =   9300
            TabIndex        =   76
            Top             =   120
            Visible         =   0   'False
            Width           =   3195
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
               TabIndex        =   80
               Top             =   800
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
               TabIndex        =   79
               Top             =   520
               Width           =   840
            End
            Begin VB.Label lblCodeDesc 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Unassigned"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   6
               Left            =   300
               TabIndex        =   78
               Top             =   1100
               Width           =   840
            End
            Begin VB.Label lblCodeDesc 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Unassigned"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   77
               Top             =   240
               Width           =   840
            End
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   6
            Left            =   5595
            TabIndex        =   31
            Top             =   5220
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   5
            Left            =   4305
            TabIndex        =   30
            Top             =   5220
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   4
            Left            =   3015
            TabIndex        =   29
            Top             =   5220
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            CausesValidation=   0   'False
            Height          =   285
            Index           =   3
            Left            =   1740
            TabIndex        =   28
            Top             =   5220
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            Height          =   285
            Left            =   1740
            TabIndex        =   8
            Tag             =   "00-Enter Position Code "
            Top             =   2565
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   5
            Tag             =   "00-Enter Status Code"
            Top             =   1568
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
            Left            =   1740
            TabIndex        =   6
            Tag             =   "EDPT-Category"
            Top             =   1900
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
            Left            =   1740
            TabIndex        =   4
            Tag             =   "00-Enter Union Code"
            Top             =   1236
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
            Left            =   1740
            TabIndex        =   2
            Tag             =   "00-Enter Location Code"
            Top             =   900
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            Height          =   285
            Left            =   1740
            TabIndex        =   1
            Tag             =   "00-Specific Department Desired"
            Top             =   572
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
            Left            =   1740
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
            Index           =   8
            Left            =   1740
            TabIndex        =   15
            Tag             =   "00-Enter Administered By Code"
            Top             =   3555
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDAB"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   9
            Left            =   1740
            TabIndex        =   17
            Tag             =   "00-Enter Section Code"
            Top             =   3892
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   7
            Left            =   1740
            TabIndex        =   14
            Tag             =   "00-Enter Region Code"
            Top             =   3228
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   1
            Left            =   3540
            TabIndex        =   12
            Tag             =   "40-Date upto and including this date forward"
            Top             =   2896
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   11
            Tag             =   "40-Date from and including this date forward"
            Top             =   2896
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   1740
            TabIndex        =   7
            Tag             =   "10-Enter Employee Number"
            Top             =   2232
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
            Left            =   1740
            TabIndex        =   18
            Tag             =   "10-Reporting Authority 1"
            Top             =   4224
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   1
            Left            =   3570
            TabIndex        =   19
            Tag             =   "10-Reporting Authority 2"
            Top             =   4224
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   2
            Left            =   5380
            TabIndex        =   20
            Tag             =   "10-Reporting Authority 3"
            Top             =   4230
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   3
            Left            =   3570
            TabIndex        =   23
            Tag             =   "40-Date upto and including this date forward"
            Top             =   4556
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   22
            Tag             =   "40-Date from and including this date forward"
            Top             =   4556
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpGrid 
            Height          =   285
            Left            =   8070
            TabIndex        =   10
            Top             =   2564
            Visible         =   0   'False
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "JBGD"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   10
            Left            =   8070
            TabIndex        =   16
            Tag             =   "00-Benefit - Group Code"
            Top             =   3560
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "BGMF"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpProv 
            Height          =   285
            Left            =   1740
            TabIndex        =   25
            Tag             =   "31-Province of Residence - Code"
            Top             =   4895
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin INFOHR_Controls.CodeLookup clpProvEmp 
            Height          =   285
            Left            =   6030
            TabIndex        =   26
            Tag             =   "31-Province of Employment - Code"
            Top             =   4890
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   11
            Left            =   10485
            TabIndex        =   34
            Tag             =   "00-Supervisory Code for cheque sorting "
            Top             =   5550
            Visible         =   0   'False
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSP"
         End
         Begin INFOHR_Controls.CodeLookup clpVadim1 
            Height          =   285
            Left            =   6030
            TabIndex        =   38
            Top             =   5550
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDV1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   12
            Left            =   8070
            TabIndex        =   13
            Top             =   2880
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   8
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   13
            Left            =   1740
            TabIndex        =   33
            Tag             =   "00-Enter Hire Code"
            Top             =   5550
            Visible         =   0   'False
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDHC"
            MaxLength       =   6
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   14
            Left            =   6960
            TabIndex        =   3
            Tag             =   "00-Enter Physical Branch Code"
            Top             =   900
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
            Left            =   6030
            TabIndex        =   37
            Tag             =   "40-Date upto and including this date forward"
            Top             =   5895
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpSenDateRange 
            Height          =   285
            Index           =   0
            Left            =   4005
            TabIndex        =   36
            Tag             =   "40-Date from and including this date forward"
            Top             =   5895
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   3
            Left            =   7200
            TabIndex        =   21
            Tag             =   "10-Reporting Authority 4"
            Top             =   4230
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.CodeLookup clpJobMaster 
            Height          =   285
            Left            =   8070
            TabIndex        =   9
            Tag             =   "01-Job code"
            Top             =   3228
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
            TabIndex        =   119
            Top             =   3285
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblToDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5640
            TabIndex        =   118
            Top             =   5940
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblSeniority 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Seniority"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3330
            TabIndex        =   117
            Top             =   5940
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblActBranch 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Physical Branch"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   116
            Top             =   951
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Hire Code"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Top             =   5602
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblSalDist 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Distribution"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6630
            TabIndex        =   108
            Top             =   2925
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblVadim1 
            AutoSize        =   -1  'True
            Caption         =   "Vadim Field 1"
            Height          =   195
            Left            =   4680
            TabIndex        =   107
            Top             =   5595
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblSupervisor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8850
            TabIndex        =   106
            Top             =   5595
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. of Residence"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   4940
            Width           =   1365
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. of Employment"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4395
            TabIndex        =   104
            Top             =   4935
            Width           =   1455
         End
         Begin VB.Label lblAnnMonth 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anniversary Month"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   5945
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lblCountry 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Country of Employment"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8760
            TabIndex        =   102
            Top             =   4935
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label lblBenGroup 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6915
            TabIndex        =   101
            Top             =   3600
            Visible         =   0   'False
            Width           =   975
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
            Left            =   6930
            TabIndex        =   99
            Top             =   2610
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblEmplStFrpmTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status From / To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   4614
            Width           =   1590
         End
         Begin VB.Label FName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   6360
            TabIndex        =   97
            Top             =   6960
            Visible         =   0   'False
            Width           =   75
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
            Left            =   7575
            TabIndex        =   96
            Top             =   4605
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblPT 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1950
            Width           =   630
         End
         Begin VB.Label lblRep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   4281
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblSection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Section"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   3948
            Width           =   540
         End
         Begin VB.Label lblAdmin 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Administered By"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   3615
            Width           =   1125
         End
         Begin VB.Label lblRegion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   3282
            Width           =   510
         End
         Begin VB.Label lblLocation 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   951
            Width           =   615
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
            TabIndex        =   89
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblLanguages 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Languages"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   5265
            Width           =   795
         End
         Begin VB.Label lblFromTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "From / To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   2949
            Width           =   1095
         End
         Begin VB.Label lblPosition 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   2616
            Width           =   975
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
            Left            =   120
            TabIndex        =   85
            Top             =   2283
            Width           =   1290
         End
         Begin VB.Label lblStatus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   1617
            Width           =   450
         End
         Begin VB.Label lblUnion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Union"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   1284
            Width           =   420
         End
         Begin VB.Label lblDept 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   618
            Width           =   825
         End
         Begin VB.Label lblDiv 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Division"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   285
            Width           =   555
         End
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   12960
      Top             =   7440
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
Attribute VB_Name = "frmRPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, SQLQ
Dim rsTABL As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkVDesc_Click()
frmDesc.Visible = chkVDesc()
End Sub

Private Sub chkWeeklyEmpList_Click()
 Call comWeeklyEmpList
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
    If FormEmplPosition% = True Then
        If Not PrtForm("Employee/Position Report Criteria", Me) Then Exit Sub
    ElseIf FormLanguages% = True Then
        If Not PrtForm("Languages Report Criteria", Me) Then Exit Sub    'laura nov 3, 1997
    Else
    End If
    
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
    Screen.MousePointer = DEFAULT
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

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
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Private Sub cmbAnnMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comCountry_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
If FormEmplPosition% = True Then    'Ticket #22144 - Hire Code
    comGroup(0).AddItem lStr("Hire Code")
End If
comGroup(0).AddItem lStr("Location")  'Jaddy jun 16,1999
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem IIf(lStr("PShift") = "PShift", "Shift", lStr("PShift"))
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem lStr("Rept. Authority 1")
If lblSalDist.Visible Then 'Ticket #21958 Franks 04/27/2012 - for Salary Distribution
    comGroup(0).AddItem lStr("Salary Distribution")
End If
comGroup(0).AddItem "(none)"

comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).AddItem "Employee Number"
comGroup(1).ListIndex = 0

comGroup(2).AddItem "Descending"
comGroup(2).AddItem "Ascending"
comGroup(2).ListIndex = 0

End Sub

Private Sub comWeeklyEmpList()
If chkWeeklyEmpList Then
    comGroup(0).Clear
    comGroup(0).AddItem "(none)"
    comGroup(0).AddItem "Employee Location"
    comGroup(0).ListIndex = 0
Else
    comGroup(0).Clear
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")  'Jaddy jun 16,1999
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If Not glbMulti Then comGroup(0).AddItem IIf(lStr("PShift") = "PShift", "Shift", lStr("PShift"))
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
End If

End Sub

Private Sub Cri_Assoc()
Dim EECri As String

If Len(clpCode(1).Text) <= 0 Then Exit Sub

If ReportSel = "POS" Then
    If glbMulti And FormEmplPosition Then
        'EECri = "{HR_JOB_HISTORY.JH_ORG} = '" & clpCode(1).Text & "' "
        'Ticket #21408
        If chkIncludeTerm Then
            EECri = "{HREMP.JH_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
        ElseIf glbMulti Then
            EECri = "{HR_JOB_HISTORY.JH_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
        End If

    Else
        'EECri = "{HREMP.ED_ORG} = '" & clpCode(1).Text & "' "
        EECri = "{HREMP.ED_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    'EECri = "ED_ORG = '" & clpCode(1).Text & "' "
    'If glbSQL Or glbOracle Then
    '    EECri = "ED_ORG in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    'Else
        EECri = "ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    'End If
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

End Sub

Private Sub Cri_Dept()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String

DeptCri = ""

If ReportSel = "POS" Then
    Call glbCri_DeptUN(clpDept.Text)
Else
    If Len(clpDept.Text) > 0 Then
        DeptCri = glbSeleDeptUn & " AND (ED_DEPTNO in ['" & Replace(clpDept.Text, ",", "','") & "'])"
    Else
        DeptCri = glbSeleDeptUn
    End If
    If Len(DeptCri) >= 1 Then
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & DeptCri
        Else
            SQLQ = DeptCri
        End If
    End If
End If

End Sub

Private Sub Cri_Div()
Dim DivCri As String

If Len(clpDiv.Text) <= 0 Then Exit Sub

If ReportSel = "POS" Then
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    Else
        glbstrSelCri = DivCri
    End If
Else
    DivCri = "(ED_DIV in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & DivCri
    Else
        SQLQ = DivCri
    End If
End If

End Sub

Private Sub Cri_Country()
Dim CountryCri As String

If UCase(comCountry.Text) = "ALL" Then Exit Sub

If Len(comCountry.Text) > 0 Then
    CountryCri = "({HREMP.ED_WORKCOUNTRY} = '" & comCountry.Text & "')"
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

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) <= 0 Then Exit Sub

If ReportSel = "POS" Then
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

End Sub

Private Sub Cri_RepAuth()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
    If glbCompSerial = "S/N - 2359W" Then 'Ticket #12156
        If FormEmplPosition% = True Then
            Exit Sub
        End If
    End If
    xTemp = False
    EECri = ""

    If Len(Trim(elpRept(0).Text)) > 0 Then
        If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
            EECri = EECri & "{HREMP.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        Else
            EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "and {HREMP.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
            Else
                EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
            End If
        Else
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "{HREMP.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
            Else
                EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
            End If
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "and {HREMP.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
            Else
                EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
            End If
        Else
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "{HREMP.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
            Else
                EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
            End If
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(3).Text)) > 0 Then
        If xTemp Then
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "and {HREMP.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
            Else
                EECri = EECri & "and {HR_JOB_HISTORY.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
            End If
        Else
            If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
                EECri = EECri & "{HREMP.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
            Else
                EECri = EECri & "{HR_JOB_HISTORY.JH_REPTAU4} = " & Trim(elpRept(3).Text) & " "
            End If
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
Dim EECri As String, LocCri As String

If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub

If ReportSel = "PRO" Then
    If Len(dlpDateRange(0).Text) > 0 Then
        'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
        '    LocCri = "(ED_LTHIRE >=" & Date_SQL(dlpDateRange(0).Text) & ")"
        'Else
            LocCri = "(ED_DOH >=" & Date_SQL(dlpDateRange(0).Text) & ")"
        'End If
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & LocCri
        Else
            SQLQ = LocCri
        End If
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
            LocCri = "(ED_LTHIRE <=" & Date_SQL(dlpDateRange(1).Text) & ")"
        'Else
            LocCri = "(ED_DOH <=" & Date_SQL(dlpDateRange(1).Text) & ")"
        'End If
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & LocCri
        Else
            SQLQ = LocCri
        End If
    End If
    Exit Sub
End If

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
    '    TempCri = "({HREMP.ED_LTHIRE} "
    'Else
        TempCri = "({HREMP.ED_DOH} "
    'End If
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(0).Text) > 0 Then    ' Daniel - 10/20/1999
    'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
    '    TempCri = "({HREMP.ED_LTHIRE} "         ' Added section to enable entering only From date, no To date.
    'Else
        TempCri = "({HREMP.ED_DOH} "         ' Added section to enable entering only From date, no To date.
    'End If
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
    GoTo Cri_FTDatst
ElseIf Len(dlpDateRange(1).Text) > 0 Then    ' Daniel - 10/20/1999
    'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
    '    TempCri = "({HREMP.ED_LTHIRE} "         ' Added section to enable entering only To date, no From date.
    'Else
        TempCri = "({HREMP.ED_DOH} "         ' Added section to enable entering only To date, no From date.
    'End If
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        'If glbCompSerial = "S/N - 2362W" Then   'Ticket #24410
        '    TempCri = "({HREMP.ED_LTHIRE}  "
        'Else
            TempCri = "({HREMP.ED_DOH}  "
        'End If
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
' Daniel - 10/20/1999 - Changed code to enable date AND other criteria simultaneously.
If Len(TempCri) > 0 Then
    If Len(glbstrSelCri) > 0 Then
      glbstrSelCri = glbstrSelCri & " AND " & TempCri
    Else
      glbstrSelCri = TempCri
    End If
End If
End Sub

Private Sub Cri_Lang1()
Dim EECri As String, OneSet%, X%
Dim strCx  As String
Dim strCa$, strC2$

OneSet% = False
strCa$ = "HR_LANGUAGE.EL_LANG_SPOKEN" ' George changed from Hremp.ED_LAng1 on Mar 21,2006.
strC2$ = "HR_LANGUAGE.EL_LANG_WRITTEN" ' George on Mar 21,2006.

For X% = 3 To 6
    If Len(clpCode(X%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then
    If glbOracle Then
        EECri = EECri & "(not isnull({" & strCa$ & "}))"
        EECri = EECri & " OR " & "(not isnull({" & strC2$ & "}))"
    Else
        EECri = EECri & "({" & strCa$ & "}<> '')"
        EECri = EECri & " OR " & "({" & strC2$ & "}) <> ''"
    End If
    EECri = EECri
    If glbstrSelCri <> "" Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
    Exit Sub
End If


For X% = 3 To 6
    If Len(clpCode(X%).Text) > 0 Then
        EECri = EECri & "({" & strCa$ & "} = '" & clpCode(X%).Text & "')"
        EECri = EECri & " OR "
    End If
Next X%

For X% = 3 To 6
    If Len(clpCode(X%).Text) > 0 Then
        EECri = EECri & "({" & strC2$ & "} = '" & clpCode(X%).Text & "')"
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            EECri = EECri & " OR "
        Else
            EECri = EECri
        End If
    End If
Next X%

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
Else
    glbstrSelCri = EECri
End If

glbiOneWhere = True

End Sub

Private Sub Cri_JobMaster() 'Ticket #27531 Franks 09/14/2015
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

Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String

If Len(clpJOB.Text) <= 0 Then Exit Sub

If glbCompSerial = "S/N - 2359W" Then 'Ticket #12156
    If FormEmplPosition% = True Then
        Exit Sub
    End If
End If

If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
    PosCri = "({HREMP.JH_JOB} = '" & clpJOB.Text & "')"
Else
    PosCri = "({HR_JOB_HISTORY.JH_JOB} = '" & clpJOB.Text & "')"
End If
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If

End Sub

Private Sub Cri_Grid()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim GirdCri As String

If Len(clpGrid.Text) <= 0 Then Exit Sub
If glbCompSerial = "S/N - 2359W" Then 'Ticket #12156
    If FormEmplPosition% = True Then
        Exit Sub
    End If
End If
If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
    GirdCri = "({HREMP.JH_GRID} = '" & clpGrid.Text & "')"
Else
    GirdCri = "({HR_JOB_HISTORY.JH_GRID} = '" & clpGrid.Text & "')"
End If
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & GirdCri
Else
    glbstrSelCri = GirdCri
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String

If Len(clpPT.Text) < 1 Then Exit Sub

If ReportSel = "POS" Then
    If (glbMulti And FormEmplPosition) Or frmRPosition.Caption = "Category/Status Report" Then
        'EECri = "{HR_JOB_HISTORY.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
        'Ticket #21408
        If (glbMulti And chkIncludeTerm) Or chkIncludeTerm Then
            EECri = "{HREMP.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
        Else
            EECri = "{HR_JOB_HISTORY.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
        End If
        
    Else
        EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
    End If
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_PT in ['" & Replace(clpPT.Text, ",", "','") & "'] "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

End Sub

Private Sub Cri_BenefitGroup()
Dim EECri As String
If Len(clpCode(10).Text) < 1 Then Exit Sub

EECri = "ED_BENEFIT_GROUP = '" & clpCode(10).Text & "' "
If Len(SQLQ) > 1 Then
    SQLQ = SQLQ & " AND " & EECri
Else
    SQLQ = EECri
End If

End Sub

Private Sub Cri_AnnMonth()
Dim EECri As String, OneSet%, X%
Dim I

If Not cmbAnnMonth.Visible Then Exit Sub
If Len(cmbAnnMonth) = 0 Then Exit Sub
I = cmbAnnMonth.ListIndex

'Ticket #23084 - City of Sarnia - use Last Hire Date.
If glbCompSerial = "S/N - 2362W" Then
    EECri = "month({HREMP.ED_LTHIRE}) = " & I
Else
    EECri = "month({HREMP.ED_DOH}) = " & I
End If

If Not glbiOneWhere Then
    glbstrSelCri = EECri
Else
    glbstrSelCri = glbstrSelCri & " AND " & EECri
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

Private Function Cri_SetAll()
Dim X%, strRName$, I

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
SQLQ = ""

' call cri models set both glbiONeWhere and strSelCri
If ReportSel = "PRO" Then
            'Laura nov 3, 1997
    Call glbCri_DeptUN(clpDept.Text)
    SQLQ = glbstrSelCri
    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
    Call Cri_Assoc
    Call Cri_Code(0)  'Jaddy jun 16,1999
    Call Cri_Code(7)  'Jaddy jun 16,1999
    Call Cri_Code(8)  'Jaddy jun 16,1999
    ' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
    Call Cri_Code(9)
    
    'Hemu - Because HR_JOB_HISTORY is not used in the report file itself. However, the Job criteria is
    'already used in the EmpWrk procedure.
    'Call Cri_Position

    Call Cri_Status
    Call Cri_PT
    Call Cri_Shift
    Call Cri_EE
    Call Cri_FTDates
    Call Cri_EmpStatFTDates
    Call Cri_ProvResidence
    Call Cri_ProvEmployment
    
    Cri_BenefitGroup
    
    'Ticket #21190
    If chkIncludeTerm Then
        Call EmpWrk
        Call EmpWrk_Terminate
    Else
        Call EmpWrk
    End If
    
    'Ticket #22682 - Delete the type of records not required by the client to show in the report
    'Call Delete_fromEmpWrk_NonShowItems
    
    X% = Cri_Sorts() 'Ticket #13495 Frank 08/09/2007
        
    glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
    
'    If glbMulti And comGroup(0).Text = lStr("Rept. Authority 1") Then
'        If Len(glbstrSelCri) >= 0 Then
'            glbstrSelCri = glbstrSelCri & " AND ({HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES') "
'        Else
'            glbstrSelCri = "({HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES')"
'        End If
'    End If
    
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
    
    'Ticket #22682 - Delete the type of records not required by the client to show in the report
    Call Delete_fromEmpWrk_NonShowItems
    
    If chkForAudit.Visible = True Then
        If chkForAudit And glbCompSerial = "S/N - 2282W" Then   'Woodbridge - Mississauga
            strRName$ = glbIHRREPORTS & "WFCrzProfil.rpt"  'Hemu - EMPHIS
        Else
            'If glbCompSerial = "S/N - 2418W" Then
             '    strRName$ = glbIHRREPORTS & "SN2418_rzProfil.rpt"
            'Else
                'Ticket #21190
                If chkIncludeTerm Then
                    strRName$ = glbIHRREPORTS & "rzProfilAT.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzProfil.rpt"
                End If
            'End If
        End If
    Else
       'If glbCompSerial = "S/N - 2418W" Then
          '       strRName$ = glbIHRREPORTS & "SN2418_rzProfil.rpt"
       ' Else
            'Ticket #21190
            If chkIncludeTerm Then
                strRName$ = glbIHRREPORTS & "rzProfilAT.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rzProfil.rpt"  'Hemu - EMPHIS
            End If
        ''End If
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
    
    'Ticket #22682 - Release 8.0
    If chkIncludeTerm Then
        Me.vbxCrystal.SubreportToChange = "RZProfilT.rpt"
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
        
        Me.vbxCrystal.Formulas(9) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
        Me.vbxCrystal.Formulas(50) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
        Me.vbxCrystal.Formulas(51) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
        Me.vbxCrystal.Formulas(52) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "
        Me.vbxCrystal.Formulas(53) = "showConDoctor = " & IIf(chkShowMedical.Value = 0, False, True) & " "
        
        Me.vbxCrystal.Formulas(58) = "showPersonalInfo = " & IIf(chkShowPersonalInfo.Value = 0, False, True) & " "
        Me.vbxCrystal.Formulas(55) = "showEmergency = " & IIf(chkShowEmergency.Value = 0, False, True) & " "
        Me.vbxCrystal.Formulas(56) = "showBanking = " & IIf(chkShowBanking.Value = 0, False, True) & " "
        Me.vbxCrystal.Formulas(57) = "showEmploymentInfo = " & IIf(chkShowEmploymentInfo.Value = 0, False, True) & " "
        
        Me.vbxCrystal.SubreportToChange = ""
    End If
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        'Changed by Frank Apr 5,2002 for the 20533 error, "cannot open database"
        'If the Databases are not in as same folder as reports are
        'For For X% = 1 To 9
        For X% = 1 To 12
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
    End If
    
    Me.vbxCrystal.WindowTitle = "Employee Profile Report"
    
    Exit Function
    
ElseIf FormEmplPosition% = True Then    'laura nov 3, 1997
        '~~~~   laura nov 3, 1997
    'Call glbCri_Dept(Me)  'laura nov 22, 1997
    Call glbCri_DeptUN(clpDept.Text)
    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
    Call Cri_Assoc
    Call Cri_Code(0)  'Jaddy jun 16,1999
    Call Cri_Code(7)  'Jaddy jun 16,1999
    Call Cri_Code(8)  'Jaddy jun 16,1999
    Call Cri_Code(9)  'Frank Nov 18,2002 for Section was being ignored
    Call Cri_Code(12) 'Ticket #21958 Franks 04/27/2012
    Call Cri_Code(13)   'Ticket #22144 - Hire Code
    If glbSamuel Then Call Cri_Code(14)    'Ticket #24162 Franks 10/03/2013 - Physical Branch
    Call Cri_Status
    Call Cri_PT
    Call Cri_Shift
    Call Cri_EE
        '~~~~~~
    Call Cri_Position
    Call Cri_Grid
    Call Cri_FTDates    'laura 03/25/98
    Call Cri_RepAuth
    Call Cri_EmpStatFTDates
    Call Cri_Country 'Ticket #16395
    Call Cri_ProvResidence
    Call Cri_ProvEmployment
    Call Cri_AnnMonth
    Call Cri_Code(11) 'Ticket #20600 Franks 09/22/2011
    Call Cri_Vadim1 'Ticket #20600 Franks 09/22/2011
        ' report name
    Call Cri_SenDates 'Ticket #24162 Franks 10/03/2013
    
    If glbWFC Then 'Ticket #27531 Franks 09/14/2015
        If FormEmplPosition% = True Then
            Call Cri_JobMaster
        End If
    End If
    
    If glbCompSerial = "S/N - 2359W" And chkWeeklyEmpList Then 'Ticket #12156
        Call Emp2359Wrk
        If comGroup(0) = "(none)" Then
            Me.vbxCrystal.Formulas(1) = "descGroup1 = ''" '"descGroup1 = '(none)'"
            Me.vbxCrystal.GroupCondition(1) = "GROUP1;{HRPARCO.PC_CO};ANYCHANGE;A"
        End If
        If comGroup(0) = "Employee Location" Then
            Me.vbxCrystal.Formulas(1) = "descGroup1 = 'Location'" '"descGroup1 = '(none)'"
            Me.vbxCrystal.GroupCondition(1) = "GROUP1;{HR_WEEKLYLIST_WRK.TT_ORG1};ANYCHANGE;A"
        End If
        strRName$ = glbIHRREPORTS & "sn23591.rpt"
        Me.vbxCrystal.WindowTitle = "Weekly Employee List Report"
        Me.vbxCrystal.ReportFileName = strRName$
        If Not (comGroup(0) = "(none)" Or comGroup(0) = "Employee Location") Then
            X% = Cri_Sorts()
        End If
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = glbstrSelCri
        End If
        Me.vbxCrystal.Connect = RptODBC_SQL 'Ticket #25890 Franks 08/15/2014
        GoTo N_Line1
    End If
    
    If frmRPosition.Caption = "Category/Status Report" Then
        If comGroup(0) = "(none)" Then
            Me.vbxCrystal.Formulas(1) = "descGroup1 = '(none)'"
            Me.vbxCrystal.GroupCondition(1) = "GROUP1;{HRPARCO.PC_CO};ANYCHANGE;A"
        End If
        strRName$ = glbIHRREPORTS & "sn2343.rpt"
        Me.vbxCrystal.WindowTitle = "Category/Status Report"
    Else
        If comGroup(0) <> "(none)" Then
            If chkIncludeTerm Then
                'strRName$ = glbIHRREPORTS & "rzpositn_T.rpt"
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "positn_T.rpt"
            Else
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "positn.rpt"
            End If
            If glbMulti Then
                Me.vbxCrystal.GroupCondition(3) = "GROUP3;{@EFullName};ANYCHANGE;A"
            End If
        Else
            If chkIncludeTerm Then
                'strRName$ = glbIHRREPORTS & "rzposit1_T.rpt"
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "posit1_T.rpt"
            Else
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "posit1.rpt"
            End If
        End If
        
        If chkLastDay And comGroup(0) <> "(none)" Then
            If chkIncludeTerm Then
                'strRName$ = glbIHRREPORTS & "rzpositL_T.rpt"
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "positL_T.rpt"
            Else
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "positL.rpt"
            End If
            If glbMulti Then
                Me.vbxCrystal.GroupCondition(3) = "GROUP3;{@EFullName};ANYCHANGE;A"
            End If
        ElseIf chkLastDay And comGroup(0) = "(none)" Then
            If chkIncludeTerm Then
                'strRName$ = glbIHRREPORTS & "rzpostL1_T.rpt"
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "postL1_T.rpt"
            Else
                strRName$ = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "postL1.rpt"
            End If
        End If
        
        'Ticket #18792
        Me.vbxCrystal.Formulas(54) = "ShowTotalOnly= " & IIf(chkTotalsOnly, 1, 0) & " "
        
        Me.vbxCrystal.WindowTitle = "Alphabetical List of Employee/Positions"
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
        ' set to sorting/grouping criteria
    X% = Cri_Sorts()   ' returns number of sections formated
        'set location for database tables
        
    'Ticket #18481 - begin
    '05/11/2010 Frank. There is no salary info on Position reports, so do not use glbNoNONE And glbNoEXEC security
    'If glbNoNONE And glbNoEXEC Then  'Hemu -EXE
    '    If Len(glbstrSelCri) >= 0 Then
    '        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
    '    Else
    '        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
    '    End If
    'ElseIf glbNoNONE Then
    '    If Len(glbstrSelCri) >= 0 Then
    '        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
    '    Else
    '        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
    '    End If
    'ElseIf glbNoEXEC Then    'Hemu -EXE
    '    If Len(glbstrSelCri) >= 0 Then
    '        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"    'Hemu -EXE
    '    Else
    '        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"   'Hemu -EXE
    '    End If
    'Else
    '    If Len(glbstrSelCri) >= 0 Then
    '        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    '    End If
    'End If
    
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
    'Ticket #18481 - end
    
    Me.vbxCrystal.Connect = RptODBC_SQL
    
            '~~~~~~~~laura nov 3, 1997   added elseif
ElseIf FormLanguages% = True Then
            'Call Cri_Dept
    'Call glbCri_Dept(Me)  'laura nov 22, 1997
    Call glbCri_DeptUN(clpDept.Text)
    Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
    Call Cri_Assoc
    Call Cri_Code(0)  'Jaddy jun 16,1999
    'Call Cri_Code(1) ' danielk - 06/20/2002 - 1 is union, handled by Cri_Assoc.
    Call Cri_Code(7)  'Jaddy jun 16,1999
    Call Cri_Code(8)  'Jaddy jun 16,1999
    Call Cri_Code(9)  'Frank Nov 18,2002 for Section was being ignored
    Call Cri_Status
    Call Cri_PT
    Call Cri_Shift
    Call Cri_EE
    Call Cri_Lang1
    
    If comGroup(0) <> "(none)" Then
        strRName$ = glbIHRREPORTS & "rzlang.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzlang1.rpt"
    End If
    Me.vbxCrystal.ReportFileName = strRName$
        
    ' set to sorting/grouping criteria
    X% = Cri_Sorts()   ' returns number of sections formated
    
    'set location for database tables

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

    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.WindowTitle = "Language Report"
    
End If

N_Line1:
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
z% = 0
X% = 0

'first set primary grouping
If comGroup(0).Text = lStr("Salary Distribution") Then 'Ticket #21958 Franks 04/27/2012
    grpField$ = "{HRSALDIST.SD_DESC}"
ElseIf comGroup(0).Text = lStr("Hire Code") Then    'Ticket #22144 - Hire Code
    grpField$ = "{tblHireCode.TB_DESC}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
End If
Y% = X% + 1

If ReportSel = "PRO" Then
    'If comGroup(0) = "(none)" Then grpField$ = "{@EFullName}"
    Me.vbxCrystal.Formulas(9) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
    Me.vbxCrystal.Formulas(50) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
    Me.vbxCrystal.Formulas(51) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
    Me.vbxCrystal.Formulas(52) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "
    Me.vbxCrystal.Formulas(53) = "showConDoctor = " & IIf(chkShowMedical.Value = 0, False, True) & " "
    Call setRptLabel(Me, 1)
Else
    If FormEmplPosition% = True Then
        If glbCompSerial = "S/N - 2362W" Then    'Ticket #24410 - City of Sarnia: Show Last Hire Date instead of DOH
            Me.vbxCrystal.Formulas(9) = "lblOHireDate='" & lStr("Last Hire Date") & "'"
        Else
            Me.vbxCrystal.Formulas(9) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
        End If
        Me.vbxCrystal.Formulas(38) = "lblPT='" & lStr("Category") & "'"
        Call setRptLabel(Me, 0)
    End If
    If comGroup(0) = "(none)" Then
        'Sorting
        GrpIdx% = comGroup(1).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        If ReportSel = "PRO" Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Else
            If chkIncludeTerm Then
                grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Else
                grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            End If
        End If
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    
        Exit Function
    End If
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(X%) = dscGroup$
End If

If comGroup(0) <> "(none)" Then
    grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(X%) = grpCond$
    
    'Sorting
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    If ReportSel = "PRO" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Else
        If chkIncludeTerm Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
    End If
    'Commented this because it gives error when Include Term selected. Also if you are getting,
    'Dll error - make sure the user has the latest info:HR script. I am using view in the report
    'that needs to be up-todate.
    'Ticket #20870 Franks 08/29/2011 - there is no group 3 for "POS" reprot
    'grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    
    Me.vbxCrystal.GroupCondition(1) = grpCond$
Else
    'Sorting
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
End If

If ReportSel = "PRO" Then
    If Left(comGroup(2), 1) = "A" Then
        Me.vbxCrystal.SortFields(0) = "+{HREMPWRK.TT_CHGDATE}"
        Me.vbxCrystal.SortFields(1) = "+{HREMPWRK.TT_PREVIEW}"
        Me.vbxCrystal.SortFields(2) = "-{HREMPWRK.TT_DATCOMP}"
        Me.vbxCrystal.SortFields(3) = "+{HREMPWRK.TT_SDATE}"
        Me.vbxCrystal.SortFields(4) = "+{HREMPWRK.TT_SEDATE}"
        Me.vbxCrystal.SortFields(5) = "-{HREMPWRK.TT_PNEXT}"
        Me.vbxCrystal.SortFields(6) = "-{HREMPWRK.TT_SKLDTE}"
        Me.vbxCrystal.SortFields(7) = "-{HREMPWRK.TT_EFDATE}"
        Me.vbxCrystal.SortFields(8) = "-{HREMPWRK.TT_YEAR}"
        Me.vbxCrystal.SortFields(9) = "-{HREMPWRK.TT_BEDATE}"
    End If
        
    Exit Function
Else
    strVis$ = "T;"
    strFVis$ = "T;"
    strPage$ = "T;"
    strSFormat$ = "GH" & CStr(Y%) & ";" & strVis$ & strPage$ & "X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF" & CStr(Y%) & ";" & strFVis$ & "X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
End If


Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
Dim EECri As String, LocCri As String
If Len(clpCode(2).Text) <= 0 Then Exit Sub

If ReportSel = "PRO" Then
    LocCri = "(ED_EMP in ['" & Replace(clpCode(2).Text, ",", "','") & "'])"
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & LocCri
    Else
        SQLQ = LocCri
    End If
Else
    If Len(clpCode(2).Text) > 0 Then
        If frmRPosition.Caption = "Category/Status Report" Then
            EECri = "{HR_JOB_HISTORY.JH_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "']"
        Else
            EECri = "{HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "'] "
        End If
    End If
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 1 Then '21July99-js-added check for other selection criteria
        'If glbiOneWhere Then         '           -commented out
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
        glbiOneWhere = True
    End If
End If
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 7 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 8 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 9 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000
    If intIdx% = 11 Then strCd$ = "HREMP.ED_SUPCODE"
    If intIdx% = 12 Then strCd$ = "HREMP.ED_SALDIST" 'Ticket #21958 Franks 04/27/2012
    If intIdx% = 13 Then strCd$ = "HREMP.ED_HIRECODE" 'Ticket #22144
    If intIdx% = 14 Then strCd$ = "HREMP.ED_SUBDEPT" 'Ticket #24162 Franks 09/19/2013

    If ReportSel = "POS" Then
            CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
        If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
            CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
        End If
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & CodeCri
        Else
            glbstrSelCri = CodeCri
        End If
    Else
        CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
        If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
            CodeCri = "(" & strCd$ & " = '" & clpDiv.Text & clpCode(intIdx%).Text & "')"
        End If
        If Len(SQLQ) > 1 Then
            SQLQ = SQLQ & " AND " & CodeCri
        Else
            SQLQ = CodeCri
        End If
    End If
End If

End Sub

Private Function CriCheck()
Dim X%, I

CriCheck = False

If Me.Caption = "Employee Profile Report" Then
    ReportSel = "PRO"
Else
    ReportSel = "POS"
End If

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
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Len(clpJOB.Text) > 0 And clpJOB.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpJOB.SetFocus
    Exit Function
End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If ReportSel = "PRO" Then
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
End If

If frmRPosition.Caption = "Employee/Position Report" Then
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
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
    If DaysBetween(dlpDateRange(2), dlpDateRange(3)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(2).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

For I = 0 To 3
    If elpRept(I).Caption = "Enter Valid Employee #" Then
        MsgBox "If Reporting Authority Entered - they must exist"
        elpRept(I).SetFocus
        Exit Function
    End If
Next

If Not elpEEID.ListChecker Then
    Exit Function
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

Private Sub Emp2359Wrk()
Dim SQLX
Dim xEmpList
Dim xDate1, xDate2
Dim rsEmp As New ADODB.Recordset
Dim rsLWRK As New ADODB.Recordset
Dim rsLJOB As New ADODB.Recordset
Dim xNum, xtot
Dim I, xJob1, xJob2, xJob3
On Error GoTo ERR_EmpWrk
gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600
If Len(dlpDateRange(0).Text) = 0 Then
  xDate1 = DateAdd("yyyy", -100, Date)
Else
  xDate1 = dlpDateRange(0).Text
End If
If Len(dlpDateRange(1).Text) = 0 Then
  'xDate2 = DateAdd("yyyy", -100, Date)     'Jaddy 10/27/99
  xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99
Else
  xDate2 = dlpDateRange(1).Text
End If
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

SQLX = "DELETE FROM HR_WEEKLYLIST_WRK " & in_SQL(glbIHRDBW)
gdbAdoIhr001.Execute SQLX
Call Pause(2)

SQLX = "SELECT ED_EMPNBR FROM HREMP "
If Len(clpJOB.Text) > 0 Or Len(clpGrid.Text) > 0 Then
    If glbOracle Then
        SQLX = SQLX & ",HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR(+)=HR_JOB_HISTORY.JH_EMPNBR "
    Else
        SQLX = SQLX & " LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR WHERE (1=1)"
    End If
    
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " AND " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    End If
    
    If Len(clpJOB.Text) > 0 Then SQLX = SQLX & " AND JH_JOB='" & clpJOB.Text & "'"
    If Len(clpGrid.Text) > 0 Then SQLX = SQLX & " AND JH_GRID='" & clpGrid.Text & "'"
    
Else
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " WHERE " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    Else
        SQLX = SQLX & " WHERE (1=1)"
    End If
End If
If glbNoNONE Then
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'NONE') "
End If
If glbNoEXEC Then
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'EXEC') "
End If
If Len(elpEEID.Text) > 0 Then
    SQLX = SQLX & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
If Len(clpDiv.Text) > 0 Then SQLX = SQLX & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If Len(clpDept.Text) > 0 Then SQLX = SQLX & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"

rsEmp.Open SQLX, gdbAdoIhr001, adOpenStatic
If rsEmp.EOF And rsEmp.BOF Then
    GoTo rr
    Exit Sub
End If
Call Pause(0.5)
rsLWRK.Open "SELECT * FROM HR_WEEKLYLIST_WRK", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
xNum = 0
If Not rsEmp.EOF Then
    xtot = rsEmp.RecordCount
End If
Do Until rsEmp.EOF
    MDIMain.panHelp(0).FloodPercent = (xNum / xtot) * 100: xNum = xNum + 1
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
    SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
    I = 1: xJob1 = "": xJob2 = "": xJob3 = ""
    If rsLJOB.State <> 0 Then rsLJOB.Close
    rsLJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsLJOB.EOF
        'If I = 1 Then '
        '    If Not IsNull(rsLJOB("JH_JOB")) Then
        '        xJob1 = GetJobCodeDesc(rsLJOB("JH_JOB")) 'GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
        '    End If
        'End If
        'If I = 2 Then
        '    If Not IsNull(rsLJOB("JH_JOB")) Then
        '        xJob2 = GetJobCodeDesc(rsLJOB("JH_JOB")) 'GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
        '    End If
        'End If
        'If I = 3 Then
        '    If Not IsNull(rsLJOB("JH_JOB")) Then
        '        xJob3 = GetJobCodeDesc(rsLJOB("JH_JOB")) 'GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
        '    End If
        'End If
        
        'Ticket #15435
        If I = 1 Then '
            If Not IsNull(rsLJOB("JH_ORG")) Then
                xJob1 = GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
            End If
        End If
        If I = 2 Then
            If Not IsNull(rsLJOB("JH_ORG")) Then
                xJob2 = GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
            End If
        End If
        If I = 3 Then
            If Not IsNull(rsLJOB("JH_ORG")) Then
                xJob3 = GetTABLDesc("EDOR", rsLJOB("JH_ORG"))
            End If
        End If
        I = I + 1
        If I > 3 Then
            GoTo exit_loop1
        End If
        rsLJOB.MoveNext
    Loop
exit_loop1:
    rsLWRK.AddNew
    rsLWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
    If Len(xJob1) > 0 Then rsLWRK("TT_ORG1") = xJob1
    If Len(xJob2) > 0 Then rsLWRK("TT_ORG2") = xJob2
    If Len(xJob3) > 0 Then rsLWRK("TT_ORG3") = xJob3
    rsLWRK.Update
    rsEmp.MoveNext
    If rsEmp.EOF Then Exit Do
Loop
rsEmp.Close
rsLWRK.Close

''xEmplist = "(" & Mid(xEmplist, 2) & ")"
'Call glbEmpWrk(xEmplist, xDate1, xDate2)
rr:
gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub
ERR_EmpWrk:
If Err = 13 Then
  FName.Visible = True
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "EMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub EmpWrk()
Dim SQLX
Dim xEmpList
Dim xDate1, xDate2
Dim rsEmp As New ADODB.Recordset

On Error GoTo ERR_EmpWrk

gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

If Len(dlpDateRange(0).Text) = 0 Then
    xDate1 = DateAdd("yyyy", -100, Date)
Else
    xDate1 = dlpDateRange(0).Text
End If

If Len(dlpDateRange(1).Text) = 0 Then
    'xDate2 = DateAdd("yyyy", -100, Date)     'Jaddy 10/27/99
    xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99
Else
    xDate2 = dlpDateRange(1).Text
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

gdbAdoIhr001.BeginTrans
SQLX = "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & "WHERE TT_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLX
gdbAdoIhr001.CommitTrans

SQLX = "SELECT ED_EMPNBR FROM HREMP "
If Len(clpJOB.Text) > 0 Or Len(clpGrid.Text) > 0 Then
    If glbOracle Then
        SQLX = SQLX & ",HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR(+)=HR_JOB_HISTORY.JH_EMPNBR "
    Else
        SQLX = SQLX & " LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR WHERE (1=1)"
    End If
    
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " AND " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    End If
    
    If Len(clpJOB.Text) > 0 Then SQLX = SQLX & " AND JH_JOB='" & clpJOB.Text & "'"
    If Len(clpGrid.Text) > 0 Then SQLX = SQLX & " AND JH_GRID='" & clpGrid.Text & "'"
    
Else
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " WHERE " & Replace(Replace(Replace(Replace(SQLQ, "{", ""), "}", ""), "[", "("), "]", ")")
    Else
        SQLX = SQLX & " WHERE (1=1)"
    End If
End If
If glbNoNONE Then
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'NONE') "
End If
If glbNoEXEC Then       'Hemu -EXE
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'EXEC') "  'Hemu -EXE
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    'Limit the user has access to these employees salary information
    If Len(glbNoAccessGrp) <> 0 Then
        SQLX = SQLX & " AND (" & glbNoAccessGrp & ") "
    End If
End If

rsEmp.Open SQLX, gdbAdoIhr001, adOpenStatic
If rsEmp.EOF And rsEmp.BOF Then
    GoTo rr
    Exit Sub
End If

MDIMain.panHelp(0).FloodPercent = 5
xEmpList = "(" & SQLX & ")"
'Do Until rsEMP.EOF
'    xEmplist = xEmplist & "," & rsEMP("ED_EMPNBR")
'    rsEMP.MoveNext
'    If rsEMP.EOF Then Exit Do
'Loop
'rsEMP.Close
'xEmplist = "(" & Mid(xEmplist, 2) & ")"
Call glbEmpWrk(xEmpList, xDate1, xDate2)

If chkForAudit Then
    If glbCompSerial = "S/N - 2282W" Then   'Woodbridge - Mississauga
        'Check if Formal Education records exists - if not then add blank
        Dim rsHREmp As New ADODB.Recordset
        Dim rsEdu As New ADODB.Recordset
        Dim rsDepend As New ADODB.Recordset
        
        SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsHREmp.EOF
            SQLQ = "SELECT EU_EMPNBR FROM HREDU WHERE EU_EMPNBR = " & rsHREmp("ED_EMPNBR")
            rsEdu.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsEdu.EOF Then
                SQLX = "INSERT INTO HREMPWRK "
                SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,TT_WRKEMP) "
                SQLX = SQLX & "VALUES ('001'," & rsHREmp("ED_EMPNBR") & ",'13','" & glbUserID & "')"
                
                gdbAdoIhr001.Execute SQLX
            End If
            rsEdu.Close
            
            SQLQ = "SELECT DP_EMPNBR FROM HRDEPEND WHERE DP_EMPNBR = " & rsHREmp("ED_EMPNBR")
            rsDepend.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsDepend.EOF Then
                SQLX = "INSERT INTO HREMPWRK "
                SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,TT_WRKEMP) "
                SQLX = SQLX & "VALUES ('001'," & rsHREmp("ED_EMPNBR") & ",'03','" & glbUserID & "')"
                
                gdbAdoIhr001.Execute SQLX
            End If
            rsDepend.Close
            
            rsHREmp.MoveNext
        Loop
        rsHREmp.Close
    End If
End If

rr:
gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub
ERR_EmpWrk:
If Err = 13 Then
  FName.Visible = True
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "EMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Form_Activate()

Call SET_UP_MODE

Screen.MousePointer = HOURGLASS

If frmRPosition.Caption = "Category/Status Report" Then 'for  "S/N - 2343W"   'ottawa ccac
    lblPT.ForeColor = &HC000C0
    lblStatus.ForeColor = &HC000C0
End If

If Me.Caption = "Employee Profile Report" Then
    lblBenGroup.Left = lblRep.Left
    lblBenGroup.Top = lblRep.Top
    lblBenGroup.Visible = True
    lblBenGroup.Alignment = vbLeftJustify
    clpCode(10).Left = elpRept(0).Left
    clpCode(10).Top = elpRept(0).Top
    clpCode(10).Visible = True
End If

'Ticket #21958 Franks 04/27/2012 - for Salary Distribution
If frmRPosition.Caption = "Employee/Position Report" Then
    lblSalDist.Caption = lStr("Salary Distribution")
    lblSalDist.Visible = True
    clpCode(12).Visible = True
    If glbSamuel Then 'Ticket #24162 Franks 10/03/2013
        Call SamuelScreenSetup
    End If
End If

Call comGrpLoad 'Ticket #21958 Franks 04/27/2012

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJOB.TransDiv = glbWFCUserSecList
End If

Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

'Ticket #26726 Franks 06/15/2015 - open it for all
'If glbWFC Then 'Ticket #25911 Franks 01/27/2015
    clpJOB.TextBoxWidth = 1315 ' 1265
'End If

'Ticket #16228 - Oxford
If glbCompSerial = "S/N - 2259W" Then
    glbMulti = True
End If

If FormLanguages% = True Then
    lblLanguages.Visible = True
    clpCode(3).Visible = True
    clpCode(4).Visible = True
    clpCode(5).Visible = True
    clpCode(6).Visible = True
    
    clpCode(3).ShowDescription = True
    clpCode(4).ShowDescription = True
    clpCode(5).ShowDescription = True
    clpCode(6).ShowDescription = True
    
    lblPosition.Visible = False
    clpJOB.Visible = False
    
    lblEmplStFrpmTo.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
    dlpDateRange(2).Visible = False
    dlpDateRange(3).Visible = False
    lblFromTo.Visible = False
    
    lblLanguages.Top = 2680
    clpCode(3).Top = 2670
    clpCode(4).Top = 2670
    clpCode(5).Top = 2670
    clpCode(6).Top = 2670
    chkVDesc.Visible = True
    chkVDesc.Top = 2665
    
    'Ticket #22682 - Release 8.0 (cleaning up)
    chkTotalsOnly.Visible = False
    frRptGrouping.Top = 9000
Else
    If FormEmplPosition% = True Then
        lblRep.Visible = True
        elpRept(0).Visible = True
        elpRept(1).Visible = True
        elpRept(2).Visible = True
        elpRept(3).Visible = True
        chkLastDay.Visible = True
        
        'Ticket #18283
        'If glbSQL And Not glbMulti Then
            chkIncludeTerm.Visible = True
        'Else
        '    chkIncludeTerm.Visible = False
        'End If
        
        'Ticket #16395 - begin
        Call addCountryItems
'        lblCountry.Top = lblLanguages.Top + 10
'        lblCountry.Left = lblLanguages.Left
'        comCountry.Top = lblLanguages.Top - 40
'        comCountry.Left = elpEEID.Left + 300
        lblCountry.Top = 4935
        lblCountry.Left = 8880
        comCountry.Top = 4875
        comCountry.Left = 10680
        lblCountry.Visible = True
        comCountry.Visible = True
        
        
        cmbAnnMonth.Top = lblTitle(1).Top
        lblAnnMonth.Top = clpCode(13).Top + 100
        
        lblTitle(1).Top = lblLanguages.Top     'Hire
        clpCode(13).Top = clpCode(3).Top
        lblVadim1.Top = lblTitle(1).Top
        clpVadim1.Top = clpCode(13).Top
        lblSupervisor.Top = lblTitle(1).Top
        clpCode(11).Top = clpCode(13).Top
        'Ticket #16395 - end
        
        'Ticket #18443
        cmbAnnMonth.Visible = True
        lblAnnMonth.Visible = True
        
        'Ticket #20600 Franks 09/22/2011
        lblSupervisor.Visible = True
        clpCode(11).Visible = True
        lblVadim1.Visible = True
        clpVadim1.Visible = True
                
        'Ticket #22144 - Hire Code
        lblTitle(1).Visible = True
        clpCode(13).Visible = True
                
        Call comAnnMonthAdding
    End If
    
    'Ticket #21190
    chkIncludeTerm.Visible = True
    
    lblLanguages.Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    If glbMultiGrid Then
        lblGrid.Visible = True
        clpGrid.Visible = True
    End If
        
    'Woodbridge - Mississauga
    If (FormLanguages% = False And FormEmplPosition% = False) And glbCompSerial = "S/N - 2282W" Then
        chkForAudit.Visible = True
    Else
        chkForAudit.Visible = False
    End If
    
    If (FormLanguages% = False And FormEmplPosition% = False) Then 'For Profile report
        'Ticket #22682 - Release 8.0
        'chkShowMedical.Visible = True  'This is part of Hide/Show Details frame now
        frHideShowDetails.Visible = True
        frHideShowDetails.Top = 6340
        chkIncludeTerm.Top = 5300   '5800
        frRptGrouping.Top = 8740
        
        chkTotalsOnly.Visible = False
        
        'Ticket #29660 - Contract Employees Enhancement
        If glbWFC Then
            chkExclCONP.Visible = True
            chkExclRET.Visible = True
            chkExclCONP.Top = 5700
            chkExclRET.Top = 5950
        Else
            chkExclCONP.Visible = False
            chkExclRET.Visible = False
        End If
    Else
        'Ticket #22682 - Release 8.0
        frHideShowDetails.Visible = False
        frRptGrouping.Top = 7440 '9000
    End If
End If
'~~~~~~~~
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
    'Call setCaption(lblShift)
    lblShift.Caption = lStr("PShift")
    If lblShift.Caption = "PShift" Then lblShift.Caption = "Shift"
End If

Call setRptCaption(Me)
lblSupervisor.Caption = lStr("Supervisor Code")
lblVadim1.Caption = lStr("Vadim Field 1")
lblSeniority.Caption = lStr("Seniority")

'Ticket #21958 Franks 04/27/2012 - moved it to Form_Activate
'Call comGrpLoad

If Me.Caption = "Employee Profile Report" Then
    lblGrp(1).Visible = True
    comGroup(2).Visible = True
Else
    lblGrp(1).Visible = False
    comGroup(2).Visible = False
End If

If glbLinamar Then clpCode(7).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbCompSerial = "S/N - 2359W" Then
    If FormEmplPosition% = True Then
        chkWeeklyEmpList.Visible = True 'Ticket #12156
    End If
End If

If glbWFC Then 'Ticket #27531 Franks 09/14/2015
    If FormEmplPosition% = True Then
        Call WFCScreenSetup
    End If
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

Call INI_Controls(Me)

panDetails.BorderStyle = 0 'no border
panWindow.BevelOuter = 0 ' no bevel

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
    panWindow.Height = Me.ScaleHeight - 200
    panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
    If panWindow.Height >= 10000 Then   '+ 230 Then
        scrControl.Value = 0
        panDetails.Top = 0
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = panWindow.Height
    End If

End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
    Resume exH

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmRPosition = Nothing  'carmen apr 2000
End Sub

Private Sub scrControl_Change()
panDetails.Top = 0 - scrControl.Value
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub

If ReportSel = "POS" Then

    EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    EECri = "ED_SHIFT = '" & txtShift.Text & "' "
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End If

glbiOneWhere = True
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
        TempCri = "({HREMP.ED_SFDATE} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    If Len(dlpDateRange(3).Text) > 0 Then
        TempCri = TempCri & "({HREMP.ED_STDATE}  "
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
Private Function getJobCodeDesc(xKey)
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String, xStr As String
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTABL.EOF Then
        xStr = rsTABL("JB_DESCR")
    End If
    rsTABL.Close
    getJobCodeDesc = Left(xStr, 20)
End Function
Private Function GetTABLDesc(xName, xKey)
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String, xStr As String
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xName & "' AND TB_KEY = '" & xKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTABL.EOF Then
        xStr = rsTABL("TB_DESC")
    End If
    rsTABL.Close
    GetTABLDesc = Left(xStr, 20)
End Function

Private Sub addCountryItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountry.AddItem Left(ctylist, X - 1)
        'comCountryOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountry.AddItem ctylist
        'comCountryOfEmp.AddItem ctylist
    End If
Loop

comCountry.ListIndex = -1
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
If InStr(xCountryList, comCountry) = 0 And comCountry <> "" Then
    xCountryList = xCountryList & "&" & comCountry
    comCountry.AddItem comCountry
    'comCountryOfEmp.AddItem comCountry
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

Private Sub comAnnMonthAdding()
'When selected by the users, the report will only show employees who have their Original Date of Hire month
'equal the Anniversary Month.
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

Private Sub EmpWrk_Terminate()
Dim SQLX
Dim xEmpList
Dim xDate1, xDate2
Dim rsEmp As New ADODB.Recordset

On Error GoTo ERR_TermEmpWrk

gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

If Len(dlpDateRange(0).Text) = 0 Then
    xDate1 = DateAdd("yyyy", -100, Date)
Else
    xDate1 = dlpDateRange(0).Text
End If

If Len(dlpDateRange(1).Text) = 0 Then
    'xDate2 = DateAdd("yyyy", -100, Date)     'Jaddy 10/27/99
    xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99
Else
    xDate2 = dlpDateRange(1).Text
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

gdbAdoIhr001.BeginTrans
SQLX = "DELETE FROM HRTERMEMPWRK " & in_SQL(glbIHRDBW) & "WHERE TT_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLX
gdbAdoIhr001.CommitTrans

SQLX = "SELECT ED_EMPNBR FROM Term_HREMP "
If Len(clpJOB.Text) > 0 Or Len(clpGrid.Text) > 0 Then
    If glbOracle Then
        SQLX = SQLX & ",Term_JOB_HISTORY WHERE Term_HREMP.ED_EMPNBR(+)=Term_JOB_HISTORY.JH_EMPNBR "
    Else
        SQLX = SQLX & " LEFT JOIN Term_JOB_HISTORY ON Term_HREMP.ED_EMPNBR=Term_JOB_HISTORY.JH_EMPNBR AND Term_HREMP.TERM_SEQ=Term_JOB_HISTORY.TERM_SEQ WHERE (1=1)"
    End If
    
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " AND " & Replace(Replace(Replace(Replace(Replace(SQLQ, "HREMP.", "Term_HREMP."), "{", ""), "}", ""), "[", "("), "]", ")")
    End If
    
    If Len(clpJOB.Text) > 0 Then SQLX = SQLX & " AND JH_JOB='" & clpJOB.Text & "'"
    If Len(clpGrid.Text) > 0 Then SQLX = SQLX & " AND JH_GRID='" & clpGrid.Text & "'"
    
Else
    If Len(SQLQ) > 1 Then
        SQLX = SQLX & " WHERE " & Replace(Replace(Replace(Replace(Replace(SQLQ, "HREMP.", "Term_HREMP."), "{", ""), "}", ""), "[", "("), "]", ")")
    Else
        SQLX = SQLX & " WHERE (1=1)"
    End If
End If
If glbNoNONE Then
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'NONE') "
End If
If glbNoEXEC Then       'Hemu -EXE
    SQLX = SQLX & " AND (ED_ORG IS NULL OR ED_ORG <> 'EXEC') "  'Hemu -EXE
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    'Limit the user has access to these employees salary information
    If Len(glbNoAccessGrp) <> 0 Then
        SQLX = SQLX & " AND (" & glbNoAccessGrp & ") "
    End If
End If

rsEmp.Open SQLX, gdbAdoIhr001X, adOpenStatic
If rsEmp.EOF And rsEmp.BOF Then
    GoTo rr
    Exit Sub
End If

MDIMain.panHelp(0).FloodPercent = 5
xEmpList = "(" & SQLX & ")"
'Do Until rsEMP.EOF
'    xEmplist = xEmplist & "," & rsEMP("ED_EMPNBR")
'    rsEMP.MoveNext
'    If rsEMP.EOF Then Exit Do
'Loop
'rsEMP.Close
'xEmplist = "(" & Mid(xEmplist, 2) & ")"
Call glbEmpWrk_Terminate(xEmpList, xDate1, xDate2)

If chkForAudit Then
    If glbCompSerial = "S/N - 2282W" Then   'Woodbridge - Mississauga
        'Check if Formal Education records exists - if not then add blank
        Dim rsHREmp As New ADODB.Recordset
        Dim rsEdu As New ADODB.Recordset
        Dim rsDepend As New ADODB.Recordset
        
        SQLQ = "SELECT ED_EMPNBR, TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsHREmp.EOF
            SQLQ = "SELECT EU_EMPNBR,TERM_SEQ FROM Term_EDU WHERE EU_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND TERM_SEQ = " & rsHREmp("TERM_SEQ")
            rsEdu.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsEdu.EOF Then
                SQLX = "INSERT INTO HRTERMEMPWRK "
                SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,TT_WRKEMP,TERM_SEQ) "
                SQLX = SQLX & "VALUES ('001'," & rsHREmp("ED_EMPNBR") & ",'13','" & glbUserID & "'," & rsHREmp("TERM_SEQ") & ")"
                
                gdbAdoIhr001X.Execute SQLX
            End If
            rsEdu.Close
            
            SQLQ = "SELECT DP_EMPNBR,TERM_SEQ FROM Term_HRDEPEND WHERE DP_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND TERM_SEQ = " & rsHREmp("TERM_SEQ")
            rsDepend.Open SQLQ, gdbAdoIhr001X, adOpenStatic
            If rsDepend.EOF Then
                SQLX = "INSERT INTO HRTERMEMPWRK "
                SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,TT_WRKEMP,TERM_SEQ) "
                SQLX = SQLX & "VALUES ('001'," & rsHREmp("ED_EMPNBR") & ",'03','" & glbUserID & "'," & rsHREmp("TERM_SEQ") & ")"
                
                gdbAdoIhr001X.Execute SQLX
            End If
            rsDepend.Close
            
            rsHREmp.MoveNext
        Loop
        rsHREmp.Close
    End If
End If

rr:
gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

ERR_TermEmpWrk:

If Err = 13 Then
    FName.Visible = True
    MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "HRTERMEMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Delete_fromEmpWrk_NonShowItems()
    'Delete respective records that client don't want to see in the report.
    Dim SQLQD As String
    
    'Hide using Suppress option in the report
    Me.vbxCrystal.Formulas(58) = "showPersonalInfo = " & IIf(chkShowPersonalInfo.Value = 0, False, True) & " "
    Me.vbxCrystal.Formulas(55) = "showEmergency = " & IIf(chkShowEmergency.Value = 0, False, True) & " "
    Me.vbxCrystal.Formulas(56) = "showBanking = " & IIf(chkShowBanking.Value = 0, False, True) & " "
    Me.vbxCrystal.Formulas(57) = "showEmploymentInfo = " & IIf(chkShowEmploymentInfo.Value = 0, False, True) & " "
    
    
'    If chkIncludeTerm Then
'        Me.vbxCrystal.SubreportToChange = "RZProfilT.rpt"
'        Me.vbxCrystal.SelectionFormula = glbstrSelCri
'
'        Me.vbxCrystal.Formulas(9) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(50) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(51) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(52) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(53) = "showConDoctor = " & IIf(chkShowMedical.Value = 0, False, True) & " "
'
'        Me.vbxCrystal.Formulas(58) = "showPersonalInfo = " & IIf(chkShowPersonalInfo.Value = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(55) = "showEmergency = " & IIf(chkShowEmergency.Value = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(56) = "showBanking = " & IIf(chkShowBanking.Value = 0, False, True) & " "
'        Me.vbxCrystal.Formulas(57) = "showEmploymentInfo = " & IIf(chkShowEmploymentInfo.Value = 0, False, True) & " "
'
'        Me.vbxCrystal.SubreportToChange = ""
'
'    End If
    
    
    'Hide by deleting the respective record from HREMPWRK
    If chkShowDependents.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '03' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '03' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowCurPosition.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '04' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '04' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowCurSalary.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '05' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '05' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowCurPerform.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '06' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '06' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    
    If chkShowPositionHist.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '07' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '07' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowSalaryHist.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '08' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '08' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowPerformHist.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '09' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '09' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    
    If chkShowLanguages.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '10' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '10' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowSkills.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '11' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '11' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowOtherEarnings.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '12' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '12' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowFormalEdu.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '13' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '13' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowCourseSeminars.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '15' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '15' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowAssociations.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '17' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '17' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowBenefits.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '19' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '19' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowBeneficiary.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '20' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '20' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
        
    If chkShowVacSickCompWSIB.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '21' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '21' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowHourlyEntitle.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '22' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '22' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowDollarEntitle.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '23' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '23' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowEmploymentHist.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '25' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '25' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
    If chkShowLeaveHist.Value = 0 Then
        SQLQD = "DELETE FROM HREMPWRK WHERE TT_RECNBR = '26' AND TT_WRKEMP = '" & glbUserID & "'"
        gdbAdoIhr001.Execute SQLQD
        
        If chkIncludeTerm Then
            SQLQD = "DELETE FROM HRTERMEMPWRK WHERE TT_RECNBR = '26' AND TT_WRKEMP = '" & glbUserID & "'"
            gdbAdoIhr001X.Execute SQLQD
        End If
    End If
    
End Sub

Private Sub SamuelScreenSetup() 'Ticket #24162 Franks 10/03/2013
    lblActBranch.Visible = True
    clpCode(14).Visible = True
    lblSeniority.Visible = True
    lblToDate.Visible = True
    dlpSenDateRange(0).Visible = True
    dlpSenDateRange(1).Visible = True
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
    lblJOB.Top = lblGrid.Top
    clpJobMaster.Top = clpGrid.Top
    lblJOB.Visible = True
    clpJobMaster.Visible = True
End Sub
