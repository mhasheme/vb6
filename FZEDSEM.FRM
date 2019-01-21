VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmREdSem 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Continuing Education Report"
   ClientHeight    =   12930
   ClientLeft      =   735
   ClientTop       =   1005
   ClientWidth     =   13800
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12930
   ScaleWidth      =   13800
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   12135
      LargeChange     =   300
      Left            =   10440
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   91
      Top             =   0
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   12975
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   22886
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
         Height          =   12735
         Left            =   120
         ScaleHeight     =   12705
         ScaleWidth      =   10065
         TabIndex        =   54
         Top             =   120
         Width           =   10095
         Begin VB.CheckBox chkExclCONP 
            Caption         =   "Exclude Employment Status of CONP"
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Tag             =   "Check to Exclude Employees with CONP Employment Status"
            Top             =   10500
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.CheckBox chkExclRET 
            Caption         =   "Exclude Employment Status of RET"
            Height          =   285
            Left            =   120
            TabIndex        =   47
            Tag             =   "Check to Exclude Employees with RET Employment Status"
            Top             =   10740
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.Frame frHideShowCourseName 
            Height          =   495
            Left            =   5760
            TabIndex        =   99
            Top             =   10200
            Visible         =   0   'False
            Width           =   4215
            Begin VB.OptionButton optHideNotShowName 
               Caption         =   "Don’t Show Course Name"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1920
               TabIndex        =   101
               Top             =   160
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton optHideShowName 
               Caption         =   "Show Course Name"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   160
               Width           =   1815
            End
         End
         Begin VB.CheckBox chkCEUCourseDescTrainer 
            Caption         =   "Show Course Description && Trainer Name"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5640
            TabIndex        =   98
            Tag             =   "00-Show Course Description & Trainer Name"
            Top             =   10245
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Frame frCourseNameDesc 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2090
            TabIndex        =   95
            Top             =   10245
            Visible         =   0   'False
            Width           =   4215
            Begin VB.OptionButton optShowCName 
               Caption         =   "Show Course Name"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   0
               TabIndex        =   97
               Top             =   0
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optShowCDesc 
               Caption         =   "Show Course Description"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2040
               TabIndex        =   96
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.CheckBox chkNoStartDate 
            Caption         =   "No Start Date"
            Height          =   285
            Left            =   5400
            TabIndex        =   23
            Tag             =   "00-Conducted By Not Equal To"
            Top             =   5130
            Width           =   1815
         End
         Begin VB.CheckBox chkNoDateComplete 
            Caption         =   "No Date Completed"
            Height          =   285
            Left            =   5400
            TabIndex        =   26
            Tag             =   "00-Conducted By Not Equal To"
            Top             =   5458
            Width           =   2055
         End
         Begin VB.TextBox txtCompanyName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2360
            MaxLength       =   50
            TabIndex        =   29
            Tag             =   "00-Company Name"
            Top             =   6470
            Width           =   4575
         End
         Begin VB.TextBox txtTrainerName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2360
            MaxLength       =   50
            TabIndex        =   30
            Tag             =   "00-Trainer Name"
            Top             =   6795
            Width           =   4575
         End
         Begin VB.TextBox txtShift 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2360
            MaxLength       =   4
            TabIndex        =   38
            Tag             =   "00-Employee Position Shift"
            Top             =   8805
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.CheckBox chkNECourseCode 
            Caption         =   "Not"
            Height          =   255
            Left            =   1400
            TabIndex        =   15
            Tag             =   "00-Course Code Not Equal To"
            Top             =   4161
            Width           =   615
         End
         Begin VB.CheckBox chkNECourseType 
            Caption         =   "Not"
            Height          =   255
            Left            =   1400
            TabIndex        =   13
            Tag             =   "00-Course Type Not Equal To"
            Top             =   3833
            Width           =   615
         End
         Begin VB.CheckBox chkNEConduct 
            Caption         =   "Not"
            Height          =   255
            Left            =   1400
            TabIndex        =   17
            Tag             =   "00-Conducted By Not Equal To"
            Top             =   4489
            Width           =   615
         End
         Begin VB.CheckBox chkNEKeyword 
            Caption         =   "Not"
            Height          =   255
            Left            =   1400
            TabIndex        =   10
            Tag             =   "00-Keyword Not Equal To"
            Top             =   3505
            Width           =   615
         End
         Begin VB.CheckBox chkNEResults 
            Caption         =   "Not"
            Height          =   255
            Left            =   1400
            TabIndex        =   33
            Tag             =   "00-Results Not Equal To"
            Top             =   7800
            Width           =   600
         End
         Begin VB.TextBox txtCourseName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2360
            MaxLength       =   125
            TabIndex        =   37
            Tag             =   "00-Enter Course Name"
            Top             =   8445
            Width           =   6850
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Tag             =   "Final sorting of records"
            Top             =   12360
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Tag             =   "Third level of grouping records"
            Top             =   12030
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Tag             =   "Second level of grouping records"
            Top             =   11700
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Tag             =   "First Level of grouping records"
            Top             =   11370
            Width           =   2325
         End
         Begin VB.OptionButton optCost 
            Caption         =   "  Cost Only"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4320
            TabIndex        =   43
            Tag             =   "Cost Only Report"
            Top             =   9945
            Width           =   1095
         End
         Begin VB.OptionButton optSum 
            Caption         =   "  Summary"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   42
            Tag             =   "Summary Report"
            Top             =   9945
            Width           =   1215
         End
         Begin VB.OptionButton optDetail 
            Caption         =   "  Detail"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2085
            TabIndex        =   41
            Tag             =   "Detail Report"
            Top             =   9945
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox txtKey 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2360
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "00-Keyword"
            Top             =   3490
            Width           =   1620
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   11
            Left            =   2040
            TabIndex        =   9
            Tag             =   "00-Enter Section Code"
            Top             =   3162
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   9
            Left            =   2040
            TabIndex        =   7
            Tag             =   "00-Enter Region Code"
            Top             =   2506
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   10
            Left            =   2040
            TabIndex        =   8
            Tag             =   "00-Enter Administered By Code"
            Top             =   2834
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDAB"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   2040
            TabIndex        =   34
            Tag             =   "00-Results Code"
            Top             =   7785
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ESRT"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   3
            Left            =   2040
            TabIndex        =   18
            Tag             =   "00-Conducted by Code"
            Top             =   4474
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ESCB"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   12
            Left            =   2040
            TabIndex        =   16
            Tag             =   "00-Course Code"
            Top             =   4146
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCD"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   2040
            TabIndex        =   14
            Tag             =   "00-Course Type Code"
            Top             =   3818
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCT"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   3
            Tag             =   "00-Enter Union Code"
            Top             =   1194
            Width           =   7155
            _ExtentX        =   12621
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
            Top             =   866
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Tag             =   "00-Specific Department Desired"
            Top             =   538
            Width           =   7155
            _ExtentX        =   12621
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
            Left            =   2040
            TabIndex        =   0
            Tag             =   "00-Specific Division Desired"
            Top             =   210
            Width           =   7155
            _ExtentX        =   12621
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
            TabIndex        =   4
            Tag             =   "00-Enter Status Code"
            Top             =   1522
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDEM"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpPT 
            Height          =   285
            Left            =   2040
            TabIndex        =   5
            Tag             =   "EDPT-Category"
            Top             =   1850
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDPT"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   25
            Tag             =   "40-Completed Date upto and including this date"
            Top             =   5458
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   24
            Tag             =   "40-Completed Date from and including this date forward"
            Top             =   5458
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpStart 
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   22
            Tag             =   "40-Start Date upto and including this date forward"
            Top             =   5130
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpStart 
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   21
            Tag             =   "40-Start Date from and including this date forward"
            Top             =   5130
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpScheduled 
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   20
            Tag             =   "40-Scheduled Date upto and including this date"
            Top             =   4802
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpScheduled 
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   19
            Tag             =   "40-Scheduled Date from and including this date forward"
            Top             =   4802
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpRenewal 
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   36
            Tag             =   "40-Renewal Date from and including this date forward"
            Top             =   8115
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpRenewal 
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   35
            Tag             =   "40-Renewal Date from and including this date forward"
            Top             =   8115
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   2040
            TabIndex        =   6
            Tag             =   "10-Enter Employee Number"
            Top             =   2178
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            TextBoxWidth    =   6835
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpJOB 
            Height          =   285
            Left            =   2040
            TabIndex        =   27
            Tag             =   "00-Enter Position Code"
            Top             =   5786
            Width           =   7155
            _ExtentX        =   12621
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
            Index           =   6
            Left            =   2040
            TabIndex        =   31
            Tag             =   "00-Co-Ordinated By"
            Top             =   7125
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCC"
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   7
            Left            =   2040
            TabIndex        =   32
            Tag             =   "00-Method Used"
            Top             =   7455
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESMU"
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   8
            Left            =   5790
            TabIndex        =   12
            Tag             =   "00-Enter CEU Type Code"
            Top             =   3495
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ESUT"
            MaxLength       =   8
         End
         Begin INFOHR_Controls.CodeLookup clpProv 
            Height          =   285
            Left            =   2040
            TabIndex        =   39
            Tag             =   "31-Province of Residence - Code"
            Top             =   9165
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin INFOHR_Controls.CodeLookup clpProvEmp 
            Height          =   285
            Left            =   2040
            TabIndex        =   40
            Tag             =   "31-Province of Employment - Code"
            Top             =   9495
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin Threed.SSOption optCEU 
            Height          =   255
            Left            =   5640
            TabIndex        =   44
            Tag             =   "Beneficiary Report"
            Top             =   9960
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "CEU Information"
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
         Begin Threed.SSOption optHideName 
            Height          =   255
            Left            =   7320
            TabIndex        =   45
            Tag             =   "Beneficiary Report"
            Top             =   9945
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Hide Employee Data"
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
         Begin INFOHR_Controls.EmployeeLookup elpReptau 
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Tag             =   "10-Reporting Authority"
            Top             =   6120
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TextBoxWidth    =   6835
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin VB.Label lblRepAuth 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rept. Authority 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   6165
            Width           =   1185
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
            TabIndex        =   94
            Top             =   9210
            Width           =   1365
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. of Employment"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   9540
            Width           =   1455
         End
         Begin VB.Label lblCEU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CEU Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   92
            Top             =   3540
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Method Used"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   90
            Top             =   7500
            Width           =   1605
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Co-Ordinated By"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   89
            Top             =   7170
            Width           =   1905
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   88
            Top             =   6515
            Width           =   2085
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Trainer Name"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   87
            Top             =   6840
            Width           =   2040
         End
         Begin VB.Label lblJobDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   9330
            TabIndex        =   86
            Top             =   5786
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   12
            Left            =   9390
            TabIndex        =   85
            Top             =   4146
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblCodeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   9390
            TabIndex        =   84
            Top             =   3818
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblJOB 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Code"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   5831
            Width           =   975
         End
         Begin VB.Label lblShift 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   8865
            Visible         =   0   'False
            Width           =   645
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
            TabIndex        =   81
            Top             =   1895
            Width           =   630
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   4191
            Width           =   1275
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
            TabIndex        =   79
            Top             =   3207
            Width           =   540
         End
         Begin VB.Label lblCourseName 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Course Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   8505
            Width           =   1215
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
            TabIndex        =   77
            Top             =   2879
            Width           =   1605
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
            TabIndex        =   76
            Top             =   2551
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
            TabIndex        =   75
            Top             =   911
            Width           =   615
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Type of Report"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   74
            Top             =   9945
            Width           =   1065
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
            TabIndex        =   73
            Top             =   12420
            Width           =   660
         End
         Begin VB.Label lblGrp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grouping #3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   12090
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
            Left            =   120
            TabIndex        =   71
            Top             =   11760
            Width           =   885
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
            TabIndex        =   70
            Top             =   11430
            Width           =   885
         End
         Begin VB.Label lblRepGrp 
            BackStyle       =   0  'Transparent
            Caption         =   "Report Grouping"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   69
            Top             =   11115
            Width           =   1575
         End
         Begin VB.Label lblSelCri 
            BackStyle       =   0  'Transparent
            Caption         =   "Selection Criteria"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblRenewal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Renewal Date Range"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   8160
            Width           =   1665
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Results"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   7830
            Width           =   765
         End
         Begin VB.Label lblEENum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   2223
            Width           =   690
         End
         Begin VB.Label lblStatus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1567
            Width           =   450
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
            TabIndex        =   63
            Top             =   583
            Width           =   825
         End
         Begin VB.Label lblRenew 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed Range"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   5503
            Width           =   1905
         End
         Begin VB.Label lblStart 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date Range"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   5175
            Width           =   1845
         End
         Begin VB.Label lblScheduled 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scheduled Date Range"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   4847
            Width           =   1920
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducted By"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   4519
            Width           =   1245
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   3863
            Width           =   1260
         End
         Begin VB.Label lblKey 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   3535
            Width           =   615
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
            TabIndex        =   56
            Top             =   1239
            Width           =   420
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
            TabIndex        =   55
            Top             =   255
            Width           =   555
         End
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   52
      Top             =   12315
      Width           =   13800
      _Version        =   65536
      _ExtentX        =   24342
      _ExtentY        =   1085
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
      BevelOuter      =   0
      Alignment       =   8
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
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
End
Attribute VB_Name = "frmREdSem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, HisSQL, HisSQL1
Dim fNEflag As Boolean
Dim fDis(26) As String
Dim xEmplist, NECri, FuncSelCri
Dim DateSelCri
Dim WrkTabSQL


Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkNEConduct_Click()
    'Ticket #22682 - Release 8.0
    If chkNEConduct.Value = vbChecked Then
        frCourseNameDesc.Visible = False
    ElseIf chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked And optSum.Value = True Then
        frCourseNameDesc.Visible = True
    End If
End Sub

Private Sub chkNEConduct_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkNECourseCode_Click()
    'Ticket #22682 - Release 8.0
    If chkNECourseCode.Value = vbChecked Then
        frCourseNameDesc.Visible = False
    ElseIf chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked And optSum.Value = True Then
        frCourseNameDesc.Visible = True
    End If
End Sub

Private Sub chkNECourseCode_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkNECourseType_Click()
    'Ticket #22682 - Release 8.0
    If chkNECourseType.Value = vbChecked Then
        frCourseNameDesc.Visible = False
    ElseIf chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked And optSum.Value = True Then
        frCourseNameDesc.Visible = True
    End If
End Sub

Private Sub chkNECourseType_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkNEKeyword_Click()
    'Ticket #22682 - Release 8.0
    If chkNEKeyword.Value = vbChecked Then
        frCourseNameDesc.Visible = False
    ElseIf chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked And optSum.Value = True Then
        frCourseNameDesc.Visible = True
    End If
End Sub

Private Sub chkNEKeyword_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkNEResults_Click()
    'Ticket #22682 - Release 8.0
    If chkNEResults.Value = vbChecked Then
        frCourseNameDesc.Visible = False
    ElseIf chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked And optSum.Value = True Then
        frCourseNameDesc.Visible = True
    End If
End Sub

Private Sub chkNEResults_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Continuing Education Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
        
    Call set_PrintState(False)
    x% = Cri_SetAll()
    ''If optHideName.Value Then 'Ticket #25221 Franks 03/19/2014
    ''    Exit Sub
    ''    MDIMain.Timer1.Enabled = False
    ''    Screen.MousePointer = DEFAULT
    ''End If
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
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "Edu/Sem", "Print")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdView_Click()
Dim x%
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
    
    x% = Cri_SetAll()
    ''If optHideName.Value Then 'Ticket #25221 Franks 03/19/2014
    ''    Exit Sub
    ''    MDIMain.Timer1.Enabled = False
    ''    Screen.MousePointer = DEFAULT
    ''End If
    
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub chkNoDateComplete_Click()
    If chkNoDateComplete Then
        dlpDateRange(0).Text = ""
        dlpDateRange(1).Text = ""
    
        dlpDateRange(0).Enabled = False
        dlpDateRange(1).Enabled = False
    Else
        dlpDateRange(0).Enabled = True
        dlpDateRange(1).Enabled = True
    End If
End Sub

Private Sub chkNoStartDate_Click()
    If chkNoStartDate Then
        dlpStart(0).Text = ""
        dlpStart(1).Text = ""
    
        dlpStart(0).Enabled = False
        dlpStart(1).Enabled = False
    Else
        dlpStart(0).Enabled = True
        dlpStart(1).Enabled = True
    End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
'If Index = 12 Then Call CourseCode_Type
End Sub

Private Sub CourseCode_Type()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String, RType, xCodeList, xValCode, xTypeList, I
Dim rsTABL As New ADODB.Recordset
On Error GoTo Dept_GL_Err

If Len(clpCode(2).Text) > 0 Then
    Exit Sub
End If
If Len(clpCode(12).Text) > 0 Then
    xCodeList = "'" & Replace(clpCode(12).Text, ",", "','") & "'"
    xCodeList = Replace(xCodeList, "'", "")
    Do While Len(xCodeList) > 0
        I = InStr(1, xCodeList, ",")
        If I = 0 Then
            xValCode = xCodeList
            xCodeList = ""
        Else
            xValCode = Mid(xCodeList, 1, I - 1)
            xCodeList = Mid(xCodeList, I + 1, Len(xCodeList) - I)
        End If
        rsTABL.Open "SELECT TB_NAME,TB_KEY,TB_USR1 FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY='" & xValCode & "'", gdbAdoIhr001
        If Not rsTABL.EOF Then
            If IsNull(rsTABL("TB_USR1")) Then
                RType = ""
            Else
                RType = rsTABL("TB_USR1")
            End If
            If Len(RType) > 0 Then
                If Len(xTypeList) = 0 Then
                    xTypeList = RType
                Else
                    xTypeList = xTypeList & "," & RType
                End If
            End If
        End If
        rsTABL.Close
    Loop
    If Len(xTypeList) > 0 Then
        Msg$ = lStr("Do you want the associated Course Type?")
        Title$ = "info:HR"
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDYES Then clpCode(2).Text = xTypeList
        'clpCode(2).Text = xTypeList
    End If
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
    ' clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Course Code Snap", "Course Code", "SELECT")
Call RollBack '21June99 js
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_CourseType()
Dim EECri As String, OneSet%, x%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False

If Len(clpCode(2).Text) = 0 Then Exit Sub
strCa$ = "HREDSEM.ES_CTYPE"

If chkNECourseType Then
    If Len(clpCode(2).Text) > 0 Then
       fNEflag = True
        NECri = NECri & " or (" & strCa$ & " IN ('" & Replace(clpCode(2).Text, ",", "','") & "'))"
    End If
'
'    For X% = 6 To 8
'      If Len(clpCode(X%).Text) > 0 Then
'        fNEflag = True
'        NECri = NECri & " or (" & strCa$ & " = '" & clpCode(X%).Text & "')"
'      End If
'    Next X%
Else
    If Len(clpCode(2).Text) > 0 Then
      EECri = EECri & "({" & strCa$ & "} IN ['" & Replace(clpCode(2).Text, ",", "','") & "'])"
      OneSet% = OneSet% - 1
      If OneSet% > 0 Then
        EECri = EECri & " OR "
      Else
        EECri = EECri
      End If
    End If
    
'    For X% = 6 To 8
'      If Len(clpCode(X%).Text) > 0 Then
'        EECri = EECri & "({" & strCa$ & "} = '" & clpCode(X%).Text & "')"
'        OneSet% = OneSet% - 1
'        If OneSet% > 0 Then
'          EECri = EECri & " OR "
'        Else
'          EECri = EECri
'        End If
'      End If
'    Next X%
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

Private Sub Cri_CourseCode()
Dim EECri As String, OneSet%, x%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False


If Not Len(clpCode(12).Text) > 0 Then Exit Sub

strCa$ = "HREDSEM.ES_CRSCODE"

If chkNECourseCode Then
        If Len(clpCode(12).Text) > 0 Then
            fNEflag = True
            NECri = NECri & " or (" & strCa$ & " in ('" & Replace(clpCode(12).Text, ",", "','") & "'))"
        End If
    
Else
      If Len(clpCode(12).Text) > 0 Then
        EECri = EECri & "({" & strCa$ & "} in ['" & Replace(clpCode(12).Text, ",", "','") & "'])"
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

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    CodeCri = ""
    If intIdx% = 0 Then fld$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then fld$ = "HREMP.ED_ORG"
    If intIdx% = 3 Then fld$ = "HREDSEM.ES_CONDUCT"
    If intIdx% = 4 Then fld$ = "HREDSEM.ES_RESULTS"
    If intIdx% = 5 Then fld$ = "HREMP.ED_EMP"
    If intIdx% = 8 Then fld$ = "HREDSEM.ES_CEUTYPE" 'Frank added Apr 13, 2007 Ticket#12859
    If intIdx% = 9 Then fld$ = "HREMP.ED_REGION"
    If intIdx% = 10 Then fld$ = "HREMP.ED_ADMINBY"
    If intIdx% = 11 Then fld$ = "HREMP.ED_SECTION"
    
    If chkNEResults And intIdx% = 4 Then
        NECri = NECri & " or (" & fld$ & " = '" & clpCode(intIdx%).Text & "')"
        fNEflag = True
    ElseIf chkNEConduct And intIdx% = 3 Then
        NECri = NECri & " or (" & fld$ & " = '" & clpCode(intIdx%).Text & "')"
        fNEflag = True
    Else
        'Ticket #22766 - Union allows multiple selection
        'If (intIdx% <> 1) And (intIdx% <> 3) And (intIdx% <> 4) And (intIdx% <> 8) And (intIdx% <> 9) And (intIdx% <> 10) And (intIdx% <> 11) Then
        If (intIdx% <> 3) And (intIdx% <> 4) And (intIdx% <> 8) And (intIdx% <> 9) And (intIdx% <> 10) And (intIdx% <> 11) Then
            CodeCri = "({" & fld$ & "} IN ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
        Else
            CodeCri = "({" & fld$ & "} = '" & clpCode(intIdx%).Text & "')"
        End If
        If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
            CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
        End If
    End If

End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    WrkTabSQL = WrkTabSQL & " AND (" & CodeCri & ")"
    glbiOneWhere = True
End If


End Sub

Private Sub Cri_DeptSQL()
Dim EECri As String
Dim countr As Integer  ' EEList_Snap is definded at form level

If optHideName.Value Then
    If Len(clpDept.Text) > 0 Then
        EECri = "({HREMP.ED_DEPTNO} in ['" & Replace(clpDept.Text, ",", "','") & "'])"
        WrkTabSQL = WrkTabSQL & " AND " & EECri & ""
    End If
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    WrkTabSQL = WrkTabSQL & " AND " & DivCri
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
    WrkTabSQL = WrkTabSQL & " AND (" & EECri & ")"
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
Dim x%, fuTempCri As String

If chkNoDateComplete Then
    TempCri = "(ISNULL({HREDSEM.ES_DATCOMP})) "
    fuTempCri = "(ES_DATCOMP IS NULL) "
    GoTo Cri_FTDatst
End If

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HREDSEM.ES_DATCOMP} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(ES_DATCOMP >=" & Date_SQL(dlpDateRange(0).Text)
    fuTempCri = fuTempCri & " AND ES_DATCOMP <= " & Date_SQL(dlpDateRange(1).Text) & ")"
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({HREDSEM.ES_DATCOMP}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
            fuTempCri = "(ES_DATCOMP >= " & Date_SQL(dlpDateRange(0).Text) & ") "
        Else
            TempCri = TempCri & " <= "
            fuTempCri = "(ES_DATCOMP <= " & Date_SQL(dlpDateRange(1).Text) & ") "
        End If
        dtYYY% = Year(dlpDateRange(x%).Text)
        dtMM% = month(dlpDateRange(x%).Text)
        dtDD% = Day(dlpDateRange(x%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next x%


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTRenewal()
'---added on Sept 24, by Laura

Dim strRenewal As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, fuTempCri As String


If Len(dlpRenewal(0).Text) > 0 And Len(dlpRenewal(1).Text) > 0 Then
    strRenewal = "({HREDSEM.ES_RENEW} "
    dtYYY% = Year(dlpRenewal(0).Text)
    dtMM% = month(dlpRenewal(0).Text)
    dtDD% = Day(dlpRenewal(0).Text)
    strRenewal = strRenewal & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpRenewal(1).Text)
    dtMM% = month(dlpRenewal(1).Text)
    dtDD% = Day(dlpRenewal(1).Text)
    strRenewal = strRenewal & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(ES_RENEW >=" & Date_SQL(dlpRenewal(0).Text)
    fuTempCri = fuTempCri & " AND ES_RENEW <= " & Date_SQL(dlpRenewal(1).Text) & ")"
    GoTo Cri_FTRen
End If

For x% = 0 To 1
    If Len(dlpRenewal(0).Text) > 0 Then
        strRenewal = "({HREDSEM.ES_RENEW}  "
        If x% = 0 Then
            strRenewal = strRenewal & " >= "
            fuTempCri = "(ES_RENEW >= " & Date_SQL(dlpRenewal(0).Text) & ") "
        Else
            strRenewal = strRenewal & " <= "
            fuTempCri = "(ES_RENEW <= " & Date_SQL(dlpRenewal(1).Text) & ") "
        End If
        dtYYY% = Year(dlpRenewal(0).Text)
        dtMM% = month(dlpRenewal(0).Text)
        dtDD% = Day(dlpRenewal(0).Text)
        strRenewal = strRenewal & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTRen
    End If
Next x%



Cri_FTRen:
If Len(strRenewal) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = strRenewal
    Else
        glbstrSelCri = glbstrSelCri & " AND " & strRenewal
    End If
    DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If
    

End Sub

Private Sub Cri_FTSched()
'---added on Sept 24, by Laura

Dim strScheduled As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, fuTempCri As String


If Len(dlpScheduled(0).Text) > 0 And Len(dlpScheduled(1).Text) > 0 Then
    strScheduled = "({HREDSEM.ES_SCHEDULED} "
    dtYYY% = Year(dlpScheduled(0).Text)
    dtMM% = month(dlpScheduled(0).Text)
    dtDD% = Day(dlpScheduled(0).Text)
    strScheduled = strScheduled & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpScheduled(1).Text)
    dtMM% = month(dlpScheduled(1).Text)
    dtDD% = Day(dlpScheduled(1).Text)
    strScheduled = strScheduled & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(ES_SCHEDULED >=" & Date_SQL(dlpScheduled(0).Text)
    fuTempCri = fuTempCri & " AND ES_SCHEDULED <= " & Date_SQL(dlpScheduled(1).Text) & ")"
    GoTo Cri_FTSch
End If

For x% = 0 To 1
    If Len(dlpScheduled(x%).Text) > 0 Then
        strScheduled = "({HREDSEM.ES_SCHEDULED}  "
        If x% = 0 Then
            strScheduled = strScheduled & " >= "
            fuTempCri = "(ES_SCHEDULED >= " & Date_SQL(dlpScheduled(0).Text) & ") "
        Else
            strScheduled = strScheduled & " <= "
            fuTempCri = "(ES_SCHEDULED <= " & Date_SQL(dlpScheduled(1).Text) & ") "
        End If
        dtYYY% = Year(dlpScheduled(x%).Text)
        dtMM% = month(dlpScheduled(x%).Text)
        dtDD% = Day(dlpScheduled(x%).Text)
        strScheduled = strScheduled & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTSch
    End If
Next x%



Cri_FTSch:
If Len(strScheduled) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = strScheduled
    Else
        glbstrSelCri = glbstrSelCri & " AND " & strScheduled
    End If
    DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTStart()
'---added on Sept 24, by Laura

Dim strStart As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, fuTempCri As String

If chkNoStartDate Then
    strStart = "(ISNULL({HREDSEM.ES_START})) "
    fuTempCri = "(ES_START IS NULL) "
    GoTo Cri_FTStr
End If

If Len(dlpStart(0).Text) > 0 And Len(dlpStart(1).Text) > 0 Then
    strStart = "({HREDSEM.ES_START} "
    dtYYY% = Year(dlpStart(0).Text)
    dtMM% = month(dlpStart(0).Text)
    dtDD% = Day(dlpStart(0).Text)
    strStart = strStart & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpStart(1).Text)
    dtMM% = month(dlpStart(1).Text)
    dtDD% = Day(dlpStart(1).Text)
    strStart = strStart & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(ES_START >=" & Date_SQL(dlpStart(0).Text)
    fuTempCri = fuTempCri & " AND ES_START <= " & Date_SQL(dlpStart(1).Text) & ")"
    GoTo Cri_FTStr
End If

For x% = 0 To 1
    If Len(dlpStart(0).Text) > 0 Then
        strStart = "({HREDSEM.ES_START}  "
        If x% = 0 Then
            strStart = strStart & " >= "
            fuTempCri = "(ES_START >= " & Date_SQL(dlpStart(0).Text) & ") "
        Else
            strStart = strStart & " <= "
            fuTempCri = "(ES_START <= " & Date_SQL(dlpStart(1).Text) & ") "
        End If
        dtYYY% = Year(dlpStart(0).Text)
        dtMM% = month(dlpStart(0).Text)
        dtDD% = Day(dlpStart(0).Text)
        strStart = strStart & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTStr
    End If
Next x%



Cri_FTStr:
If Len(strStart) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = strStart
    Else
        glbstrSelCri = glbstrSelCri & " AND " & strStart
    End If
    DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Key()
Dim strKey As String

    If Len(txtKey.Text) < 1 Then Exit Sub
    
    'strKey = "{HREDSEM.ES_KEYWORD}= '" & txtKey.Text & "'"
    
    If chkNEKeyword Then
'        strKey = "(ISNULL({HREDSEM.ES_KEYWORD}) OR Not ({HREDSEM.ES_KEYWORD}= '" & txtKey.Text & "'))"
        If glbSQL Or glbOracle Then
            NECri = NECri & " or (UPPER(HREDSEM.ES_KEYWORD) = '" & UCase(txtKey) & "')"
        Else
            NECri = NECri & " or (UCASE(HREDSEM.ES_KEYWORD) = '" & UCase(txtKey) & "')"
        End If
        fNEflag = True
    Else
        strKey = "UPPERCASE({HREDSEM.ES_KEYWORD})= '" & UCase(txtKey.Text) & "'"
    End If
    If Len(strKey) >= 0 Then
        If glbiOneWhere Then
            glbstrSelCri = glbstrSelCri & " AND " & strKey
        Else
            glbstrSelCri = strKey
        End If
    glbiOneWhere = True
    End If
End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
WrkTabSQL = WrkTabSQL & " AND (" & EECri & ")"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Job()
Dim EECri As String, OneSet%, x%
Dim strCa$

If Len(clpJOB.Text) < 1 Then Exit Sub

'EECri = "{HRJOB.ED_PT}= '" & clpJOB.Text & "'"
strCa$ = "HR_JOB_HISTORY.JH_JOB"

If Len(clpJOB.Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} in ['" & Replace(clpJOB.Text, ",", "','") & "']) AND ({HR_JOB_HISTORY.JH_CURRENT}) "
    OneSet% = OneSet% - 1
    If OneSet% > 0 Then
        EECri = EECri & " OR "
    Else
        EECri = EECri
    End If
End If


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    'Ticket #27539 - S.U.C.C.E.S.S. wants to use Reporting Authority 2 instead of 1.
    'If glbCompSerial = "S/N - 2422W" Then
    '    EECri = "{HR_JOB_HISTORY.JH_REPTAU2} IN [" & getEmpnbr(elpReptau.Text) & "] "
    'Else
        EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpReptau.Text) & "] AND ({HR_JOB_HISTORY.JH_CURRENT}) "
    'End If
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

Private Sub Cri_CourseName()
Dim EECri As String, OneSet%, x%

If Len(txtCourseName.Text) < 1 Then Exit Sub

EECri = "Instr(1, UpperCase({HREDSEM.ES_COURSE}), '" & UCase(txtCourseName.Text) & "') > 0"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ProvResidence()
Dim EECri As String, OneSet%, x%

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
Dim EECri As String, OneSet%, x%

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
Dim x%
Dim SQLQ
Dim SQLR
Cri_SetAll = False
Dim rsES As New ADODB.Recordset
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
FuncSelCri = ""
DateSelCri = ""
WrkTabSQL = "" 'Ticket #25221 Franks 03/19/2014
' call cri models set both glbiONeWhere and strSelCri

NECri = "0<>0"

' For NE (Not equal)
fNEflag = False

Call Cri_Code(3)
Call Cri_Code(4)
Call Cri_Code(8)
Call Cri_CourseType     'added on Sept 25, by Laura
Call Cri_CourseCode
Call Cri_Key        'added on Sept 24, by Laura
Call Cri_CourseName

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

FuncSelCri = glbstrSelCri
If Len(FuncSelCri) > 0 Then WrkTabSQL = glbstrSelCri Else WrkTabSQL = "(1=1)"
Call Cri_FTDates
Call Cri_FTSched    'added on Sept 24, by Laura
Call Cri_FTStart    'added on Sept 24, by Laura
Call Cri_FTRenewal  'added on Sept 24, by Laura
Call Cri_CompanyName
Call Cri_Trainer
Call Cri_Coord
Call Cri_Method

If fNEflag Then
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.CommitTrans
    SQLQ = "INSERT INTO HR_EMPLIST_WRK (TT_COMPNO,TT_EMPNBR,TT_WRKEMP)  "
    SQLQ = SQLQ & in_SQL(glbIHRDBW) & ""
    SQLQ = SQLQ & "SELECT '001',ED_EMPNBR,'" & glbUserID & "' FROM HREMP WHERE 0=0 "
    
    'Franks Jul 19,02 begin #2581
    'If Len(NECri) > 0 Then
    '    SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & NECri
    '    If Len(DateSelCri) > 0 Then SQLQ = SQLQ & " AND (" & DateSelCri & ")"
    '    SQLQ = SQLQ & " ) "
    'End If
    'SQLR = Replace(Replace(FuncSelCri, "}", ""), "{", "")
    'If Len(SQLR) > 0 Then
    '    SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & SQLR
    '    If Len(DateSelCri) > 0 Then SQLQ = SQLQ & " AND (" & DateSelCri & ")"
    '    SQLQ = SQLQ & " ) "
    'End If
    'gdbAdoIhr001.BeginTrans
    'gdbAdoIhr001.Execute SQLQ
    'gdbAdoIhr001.CommitTrans
 
    SQLR = DateSelCri
    If Len(SQLR) = 0 Then
        If Len(NECri) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & NECri & ") "
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    Else
        'Ticket# 7301 Comment by Frank. Add all employees here, then delete the NE employees later
        'If Len(SQLR) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & SQLR & ") "
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
        'Frank 07/15/04 Ticket# 6546
        'SQLQ = "DELETE FROM HR_EMPLIST_WRK IN '" & glbIHRDBW & "' " & "WHERE TT_EMPNBR IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & NECri & " AND " & SQLR & ") "
        SQLQ = "DELETE FROM HR_EMPLIST_WRK " & in_SQL(glbIHRDBW) & " WHERE TT_EMPNBR IN (SELECT ES_EMPNBR FROM HREDSEM WHERE " & NECri & " AND " & SQLR & ") "
        SQLQ = SQLQ & " AND TT_WRKEMP='" & glbUserID & "'"
        
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    'Franks Jul 19,02 End
    
    glbstrSelCri = "{HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "' "
    glbiOneWhere = True
End If

Call glbCri_DeptUN(clpDept.Text)
Call Cri_DeptSQL '
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Code(0)
Call Cri_Code(1)
Call Cri_Code(5)
Call Cri_Code(9)
Call Cri_Code(10)
Call Cri_Code(11)
Call Cri_Job
Call Cri_Sup
Call Cri_Trainer
Call Cri_ProvResidence
Call Cri_ProvEmployment

'Ticket #25306 - Jerry said this for everyone
'If glbWFC Then
    If optHideName.Value Then 'Ticket #25221 Franks 03/20/2014
        Call WFCHideNameRpt
        If glbWFC And optHideShowName.Value Then 'Ticket #24708 Franks 04/08/2015
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzeseNoEmpN.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzeseNoEmp.rpt"
        End If
        Me.vbxCrystal.SelectionFormula = "{HREDWRK.ES_WRKEMP} = '" & glbUserID & "'"
        Me.vbxCrystal.Formulas(1) = "descGroup1 = '" & comGroup(0).Text & ":'"
        If comGroup(0) <> "(none)" Then
            Me.vbxCrystal.Formulas(2) = "showgroup1=true"
        Else
            Me.vbxCrystal.Formulas(2) = "showgroup1=false"
        End If
        If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
            Me.vbxCrystal.Formulas(3) = "SeleDaterange='Criteria: Date Completed Range - From " & CVDate(dlpDateRange(0).Text) & " to " & CVDate(dlpDateRange(1).Text) & "' "
        End If
        Me.vbxCrystal.Connect = RptODBC_SQL
        Exit Function
    End If
'End If

' report name
If optDetail.Value = True Then      'if added on Oct 10 by Laura
    If comGroup(0) <> "(none)" Then
        If glbCompSerial = "S/N - 2214W" Then    'jgr 6/5/98 To call Casey reports.
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "casey21.rpt"
        ElseIf glbCompSerial = "S/N - 2188W" Then  'Frank 4/13/07 for Chatham-Kent Ticket #12859
            Call setRptLabelEdSem
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2188esem.rpt"
        ElseIf glbCompSerial = "S/N - 2377W" And Not fNEflag Then  'Hemu - Town of Grimsby - Ticket #16766
            setRptLabelEdSem_More
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2377_rzesem.rpt"
        Else
            Call setRptLabelEdSem
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem.rpt"
        End If
    Else
        If glbCompSerial = "S/N - 2214W" Then    'jgr 6/5/98 To call Casey reports.
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "casey22.rpt"
        ElseIf glbCompSerial = "S/N - 2188W" Then  'Frank 4/13/07 for Chatham-Kent Ticket #12859
            Call setRptLabelEdSem
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2188esem1.rpt"
        ElseIf glbCompSerial = "S/N - 2377W" And Not fNEflag Then  'Hemu - Town of Grimsby - Ticket #16766
            setRptLabelEdSem_More
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2377_rzesem1.rpt"
        Else
            Call setRptLabelEdSem
            Call setRptLabelEdSemMore
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1.rpt"
        End If
    End If
ElseIf optSum.Value = True Then
    If comGroup(0) <> "(none)" Then
        Call setRptLabelEdSem
        If glbCompSerial = "S/N - 2380W" Then   'Ticket #22733 - VitalAire
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2380_ContEduS.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemS.rpt"
        
            'Ticket #22682 - Release 8.0
            If optShowCName.Value = True And fNEflag Then
                Me.vbxCrystal.Formulas(50) = "showCNameDesc='Name'"
            ElseIf optShowCDesc.Value = True Then
                Me.vbxCrystal.Formulas(50) = "showCNameDesc='Desc'"
            End If
        End If
    Else
        Call setRptLabelEdSem
        Call setRptLabelEdSemMore
        If glbCompSerial = "S/N - 2380W" Then   'Ticket #22733 - VitalAire
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2380_ContEdu1S.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1S.rpt"
        
            'Ticket #22682 - Release 8.0
            If optShowCName.Value = True And Not fNEflag Then
                Me.vbxCrystal.Formulas(50) = "showCNameDesc='Name'"
            ElseIf optShowCDesc.Value = True Then
                Me.vbxCrystal.Formulas(50) = "showCNameDesc='Desc'"
            End If
        End If
    End If
ElseIf optCEU.Value = True Then
    'Ticket #24157 - Municipality of North Perth
    vbxCrystal.Formulas(51) = "lblCompleted='" & lStr("Date Completed") & "'"
    If comGroup(0) <> "(none)" Then
        If chkCEUCourseDescTrainer.Value = vbChecked Then
            vbxCrystal.Formulas(52) = "lblTrainerName='" & lStr("Trainer Name") & "'"
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceuT.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu.rpt"
        End If
    Else
        If chkCEUCourseDescTrainer.Value = vbChecked Then
            vbxCrystal.Formulas(52) = "lblTrainerName='" & lStr("Trainer Name") & "'"
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1T.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1.rpt"
        End If
    End If
Else    'added on Oct 10 by Laura
    If comGroup(0) <> "(none)" Then
        If glbCompSerial = "S/N - 2214W" Then    'jgr 6/5/98 To call Casey reports.
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "casey23.rpt"
        Else
            Call setRptLabelEdSem
            Call setRptLabelEdSemMore
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemc.rpt"
        End If
    Else
        If glbCompSerial = "S/N - 2214W" Then    'jgr 6/5/98 To call Casey reports.
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "casey24.rpt"
        Else
            Call setRptLabelEdSem
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemc1.rpt"
        End If
    End If
End If

If fNEflag Then 'Any NE checkbox selected
    If comGroup(0) <> "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesene.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesene1.rpt"
    End If
End If

If glbCompSerial = "S/N - 2377W" And Not fNEflag And (InStr(Me.vbxCrystal.ReportFileName, "SN2377_rzesem.rpt") > 0 Or InStr(Me.vbxCrystal.ReportFileName, "SN2377_rzesem1.rpt") > 0) Then 'Hemu - Town of Grimsby - Ticket #16766
Else
    x% = Cri_Sorts()   ' returns number of sections formated
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

'Ticket #28040 - Grouping #1 = Reporting Authority 1 - Show based on Current / Multi Position
If comGroup(0).Text = lStr("Rept. Authority 1") Then
    If Len(glbstrSelCri) >= 0 Then
        If glbMulti Then
            glbstrSelCri = glbstrSelCri & " AND {HR_JOB_HISTORY.JH_CURRENT} = True AND {HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES'"
        Else
            glbstrSelCri = glbstrSelCri & " AND {HR_JOB_HISTORY.JH_CURRENT} = True"
        End If
    Else
        If glbMulti Then
            glbstrSelCri = "{HR_JOB_HISTORY.JH_CURRENT} = True AND {HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES'"
        Else
            glbstrSelCri = "{HR_JOB_HISTORY.JH_CURRENT} = True"
        End If
    End If
End If

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    If Not fNEflag Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 8
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
    End If
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Continuing Education Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Education Report", "Select")
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
Dim grpCond$, grpField$, I
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
' first set primary grouping
Y% = 0
If fNEflag Then
    'comGroup(1).ListIndex = 2
    For I = 1 To 26
        Me.vbxCrystal.Formulas(I + 2) = (fDis(I))
    Next I
End If
grpField$ = getEGroup(comGroup(0).Text)
If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$

    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    strSFormat$ = "GH1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = IIf(fNEflag, "(none)", "{tblCourse.TB_DESC}")
        Case 2: grpField$ = IIf(fNEflag, "(none)", "{tblCourseCode.TB_DESC}") 'Ticket #25148 Franks 03/04/2014
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        If comGroup(1).ListIndex <> 0 Then
            'Me.vbxCrystal.SectionFormat(0) = "GF4;F;X;X;X;X;X;X"
            'Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
        End If
        If Not fNEflag And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "SN2380_ContEduS.rpt" Then 'And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesemceu.rpt" And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesemceuT.rpt" Then   'Ticket #23862 - VitalAire, Ticket #24157 - Municipality of North Perth
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{tblCourse.TB_DESC}"
                Case 1: grpField$ = "{tblCourseCode.TB_DESC}" 'Ticket #25148 Franks 03/04/2014
                Case 2: grpField$ = "(none)"
            End Select
            If grpField$ <> "(none)" Then
                grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(2) = grpCond$
            Else
                grpCond$ = "GROUP" & CStr(3) & ";" & "{HREDSEM.ES_DATCOMP}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(2) = grpCond$
            End If
        End If
    'ElseIf Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu.rpt" Then
    '    Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
    '    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = IIf(fNEflag, "(none)", "{tblCourse.TB_DESC}")
        Case 2: grpField$ = IIf(fNEflag, "(none)", "{tblCourseCode.TB_DESC}") 'Ticket #25148 Franks 03/04/2014
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        'dscGroup$ = comGroup(1).Text               'laura 03/31/98
        'dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        'Me.vbxCrystal.Formulas(0) = dscGroup$
        '~~~
        If Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesem1S.rpt" And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesem1.rpt" Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
            
            Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
            Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
            Me.vbxCrystal.SectionFormat(3) = "GF3;F;X;F;X;X;X;X"
            Me.vbxCrystal.SectionFormat(4) = "GF4;F;X;F;X;X;X;X"
        End If
    Else
        If Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1.rpt" Or Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1T.rpt" Then     'Ticket #24157 - Municipality of North Perth
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{tblCourse.TB_DESC}"
                Case 1: grpField$ = "{tblCourseCode.TB_DESC}" 'Ticket #25148 Franks 03/04/2014
                Case 2: grpField$ = "(none)"
            End Select
            If grpField$ <> "(none)" Then
                grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
            Else
                grpCond$ = "GROUP" & CStr(1) & ";" & "{HREDSEM.ES_DATCOMP}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
            
                Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
                Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
            End If
        ElseIf Not fNEflag And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "SN2380_ContEdu1S.rpt" And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesemceu1.rpt" And Me.vbxCrystal.ReportFileName <> glbIHRREPORTS & "rzesemceu1T.rpt" Then   'Ticket #23862 - VitalAire, Ticket #24157 - Municipality of North Perth
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{tblCourse.TB_DESC}"
                Case 1: grpField$ = "{tblCourseCode.TB_DESC}" 'Ticket #25148 Franks 03/04/2014
                Case 2: grpField$ = "(none)"
            End Select
            If grpField$ <> "(none)" Then
                If Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1S.rpt" Or Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1.rpt" Then
                    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
                    Me.vbxCrystal.GroupCondition(1) = grpCond$
                    
                    Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
                    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
                    Me.vbxCrystal.SectionFormat(3) = "GF3;F;X;F;X;X;X;X"
                    Me.vbxCrystal.SectionFormat(4) = "GF4;F;X;F;X;X;X;X"
                Else
                    grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
                    Me.vbxCrystal.GroupCondition(2) = grpCond$
                End If
            Else
                grpCond$ = "GROUP" & CStr(1) & ";" & "{HREDSEM.ES_DATCOMP}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
                
                Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
                Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
                Me.vbxCrystal.SectionFormat(3) = "GF3;F;X;F;X;X;X;X"
                Me.vbxCrystal.SectionFormat(4) = "GF4;F;X;F;X;X;X;X"
                
                If Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1S.rpt" Or Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesem1.rpt" Then
                    Me.vbxCrystal.SectionFormat(5) = "GH2;X;F;X;X;X;X;X"
                    Me.vbxCrystal.SectionFormat(6) = "GF2;F;X;F;X;X;X;X"
                End If
            End If
        ElseIf Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1.rpt" Or Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzesemceu1T.rpt" Then     'Ticket #24157 - Municipality of North Perth
            Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
            Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
        End If
    End If
'    Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
'    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
'    Me.vbxCrystal.SectionFormat(3) = "GF3;F;X;F;X;X;X;X"
'    Me.vbxCrystal.SectionFormat(4) = "GF4;F;X;F;X;X;X;X"
    
End If


Cri_Sorts = z% ' next section number to format

End Function

Private Sub GetDisData()
Dim I As Integer
Dim listCourseTypes
Dim listCourseCodes
Dim listJobs
    For I = 1 To 26
        fDis(I) = ""
    Next I
    If Len(clpDiv.Text) > 0 Then
        fDis(1) = "Dis(1) = '" & lStr(lblDiv) & " = ' + chr(34) + '" & Replace(clpDiv.Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpDept.Text) > 0 Then
        fDis(2) = "Dis(2) = '" & lStr(lblDept) & " = ' + chr(34) + '" & Replace(clpDept.Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpCode(0).Text) > 0 Then
        fDis(3) = "Dis(3) = '" & lStr(lblLocation) & " = ' + chr(34) + '" & Replace(clpCode(0).Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpCode(1).Text) > 0 Then
        fDis(4) = "Dis(4) = '" & lStr(lblUnion) & " = ' + chr(34) + '" & Replace(clpCode(1).Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpCode(5).Text) > 0 Then
        fDis(5) = "Dis(5) = '" & lStr(lblStatus) & " = ' + chr(34) + '" & Replace(clpCode(5).Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpPT.Text) > 0 Then
        fDis(6) = "Dis(6) = '" & lStr(lblPT) & " = ' + chr(34) + '" & Replace(clpPT.Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(elpEEID.Text) > 0 Then
        fDis(7) = "Dis(7) = '" & lblEENum(0) & " = ' + chr(34) + '" & Replace(elpEEID.Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(txtKey) > 0 Then
        If chkNEKeyword Then
        fDis(8) = "Dis(8) = '" & lStr(lblKey) & " not equal to ' + chr(34) + '" & Replace(txtKey, "'", "''") & "' + chr(34)"
        Else
        fDis(8) = "Dis(8) = '" & lStr(lblKey) & " = ' + chr(34) + '" & Replace(txtKey, "'", "''") & "' + chr(34)"
        End If
    End If
    If Len(clpCode(2).Text) > 0 Then
        fDis(9) = "Dis(9) ="
        listCourseTypes = Split(clpCode(2).Text, ",")
        If chkNECourseType Then
            fDis(9) = fDis(9) & " '" & lStr(lblBCode(1).Caption) & " not equal to "
        Else
            fDis(9) = fDis(9) & " '" & lStr(lblBCode(1).Caption) & " = "
        End If
        For I = 0 To UBound(listCourseTypes)
            lblCodeDesc(2) = getCourseDesc("ESCT", listCourseTypes(I))
            'If chkNECourseType Then
            '    fDis(9) = fDis(9) & " '" & lblBCode(1).Caption & " not equal to ' + chr(34) + '" & Replace(lblCodeDesc(2), "'", "''") & "' + chr(34)"
            'Else
            '    fDis(9) = fDis(9) & " '" & lblBCode(1).Caption & " = ' + chr(34) + '" & Replace(lblCodeDesc(2), "'", "''") & "' + chr(34)"
            'End If
            'If I <> UBound(listCourseTypes) Then fDis(9) = fDis(9) & " + CHR(10) +CHR(13)+ "
            If chkNECourseType Then
                fDis(9) = fDis(9) & " ' + chr(34) + '" & Replace(lblCodeDesc(2), "'", "''") & "' + chr(34)"
            Else
                fDis(9) = fDis(9) & " ' + chr(34) + '" & Replace(lblCodeDesc(2), "'", "''") & "' + chr(34)"
            End If
            If I <> UBound(listCourseTypes) Then fDis(9) = fDis(9) & " + ', "
        Next
    End If
    
    If Len(clpCode(12).Text) > 0 Then
        listCourseCodes = Split(clpCode(12).Text, ",")
        fDis(13) = "Dis(13) ="
        If chkNECourseCode Then
            fDis(13) = fDis(13) & " '" & lStr(lblBCode(0)) & " not equal to "
        Else
            fDis(13) = fDis(13) & " '" & lStr(lblBCode(0)) & " = "
        End If
        For I = 0 To UBound(listCourseCodes)
            lblCodeDesc(12) = getCourseDesc("ESCD", listCourseCodes(I))
            'If chkNECourseCode Then
            '    fDis(13) = fDis(13) & " '" & lblBCode(0) & " not equal to ' + chr(34) + '" & Replace(lblCodeDesc(12), "'", "''") & "' + chr(34)"
            'Else
            '    fDis(13) = fDis(13) & " '" & lblBCode(0) & " = ' + chr(34) + '" & Replace(lblCodeDesc(12), "'", "''") & "' + chr(34)"
            'End If
            'If I <> UBound(listCourseCodes) Then fDis(13) = fDis(13) & " +CHR(10) +CHR(13)+ "
            If chkNECourseCode Then
                fDis(13) = fDis(13) & " ' + chr(34) + '" & Replace(lblCodeDesc(12), "'", "''") & "' + chr(34)"
            Else
                fDis(13) = fDis(13) & " ' + chr(34) + '" & Replace(lblCodeDesc(12), "'", "''") & "' + chr(34)"
            End If
            If I <> UBound(listCourseCodes) Then fDis(13) = fDis(13) & " + ', "
        Next
    End If
    
    If Len(clpJOB.Text) > 0 Then
        fDis(14) = "Dis(14) ="
        listJobs = Split(clpJOB.Text, ",")
        If chkNECourseType Then
            fDis(14) = fDis(14) & " '" & lblJOB.Caption & " not equal to "
        Else
            fDis(14) = fDis(14) & " '" & lblJOB.Caption & " = "
        End If
        For I = 0 To UBound(listJobs)
            lblJobDesc = getCourseDesc("JOB", listJobs(I))
            'If chkNECourseType Then
            '    fDis(14) = fDis(14) & " '" & lblJOB.Caption & " not equal to ' + chr(34) + '" & Replace(lblJobDesc, "'", "''") & "' + chr(34)"
            'Else
            '    fDis(14) = fDis(14) & " '" & lblJOB.Caption & " = ' + chr(34) + '" & Replace(lblJobDesc, "'", "''") & "' + chr(34)"
            'End If
            'If I <> UBound(listJobs) Then fDis(14) = fDis(14) & " + CHR(10) +CHR(13)+ "
            If chkNECourseType Then
                fDis(14) = fDis(14) & " ' + chr(34) + '" & Replace(lblJobDesc, "'", "''") & "' + chr(34)"
            Else
                fDis(14) = fDis(14) & " ' + chr(34) + '" & Replace(lblJobDesc, "'", "''") & "' + chr(34)"
            End If
            If I <> UBound(listJobs) Then fDis(14) = fDis(14) & " + ', "
        Next
    End If

    If Len(clpCode(3).Text) > 0 Then
        If chkNEConduct Then
        fDis(17) = "Dis(17) = '" & lStr(lblBCode(2)) & " not equal to ' + chr(34) + '" & Replace(clpCode(3).Caption, "'", "''") & "' + chr(34)"
        Else
        fDis(17) = "Dis(17) = '" & lStr(lblBCode(2)) & " = ' + chr(34) + '" & Replace(clpCode(3).Caption, "'", "''") & "' + chr(34)"
        End If
    End If
    If Len(dlpScheduled(0).Text) > 0 Then
        fDis(18) = "Dis(18) = '" & lStr(lblScheduled) & " >= ' + chr(34) + '" & Replace(dlpScheduled(0).Text, "'", "''") & "' + chr(34)"
        If Len(dlpScheduled(1).Text) > 0 Then
            fDis(18) = fDis(18) & " + ' And <= ' + chr(34) + '" & Replace(dlpScheduled(1).Text, "'", "''") & "' + chr(34)"
        End If
    End If
    If Len(dlpStart(0).Text) > 0 Then
        fDis(19) = "Dis(19) = '" & lStr(lblStart) & " >= ' + chr(34) + '" & Replace(dlpStart(0).Text, "'", "''") & "' + chr(34)"
        If Len(dlpStart(1).Text) > 0 Then
            fDis(19) = fDis(19) & " + ' And <= ' + chr(34) + '" & Replace(dlpStart(1).Text, "'", "''") & "' + chr(34)"
        End If
    End If
    If Len(dlpDateRange(0).Text) > 0 Then
        fDis(20) = "Dis(20) = '" & lStr(lblRenew) & " >= ' + chr(34) + '" & dlpDateRange(0).Text & "' + chr(34)"
        If Len(dlpDateRange(1).Text) > 0 Then
            fDis(20) = fDis(20) & " + ' And <= ' + chr(34) + '" & dlpDateRange(1).Text & "' + chr(34)"
        End If
    End If
    If Len(clpCode(9).Text) > 0 Then
        fDis(21) = "Dis(21) = '" & lStr(lblRegion) & " = ' + chr(34) + '" & clpCode(9).Caption & "' + chr(34)"
    End If
    If Len(clpCode(11).Text) > 0 Then
        fDis(22) = "Dis(22) = '" & lStr(lblSection) & " = ' + chr(34) + '" & clpCode(11).Caption & "' + chr(34)"
    End If
    If Len(clpCode(4).Text) > 0 Then
        If chkNEResults Then
        fDis(23) = "Dis(23) = '" & lStr(lblBCode(3)) & " not equal to ' + chr(34) + '" & clpCode(4).Caption & "' + chr(34)"
        Else
        fDis(23) = "Dis(23) = '" & lStr(lblBCode(3)) & " = ' + chr(34) + '" & clpCode(4).Caption & "' + chr(34)"
        End If
    End If
    If Len(dlpRenewal(0).Text) > 0 Then
        fDis(24) = "Dis(24) = '" & lStr(lblRenewal) & " >= ' + chr(34) + '" & dlpRenewal(0).Text & "' + chr(34)"
        If Len(dlpScheduled(1).Text) > 0 Then
            fDis(24) = fDis(24) & " + ' And <= ' + chr(34) + '" & dlpRenewal(1).Text & "' + chr(34)"
        End If
    End If
    If Len(clpCode(10).Text) > 0 Then
        fDis(25) = "Dis(25) = '" & lStr(lblAdmin) & " = ' + chr(34) + '" & clpCode(10).Caption & "' + chr(34)"
    End If
    If Len(txtCourseName) > 0 Then
        fDis(26) = "Dis(26) = '" & lStr(lblCourseName) & " = ' + chr(34) + '" & Replace(txtCourseName, "'", "''") & "' + chr(34)"
    End If
    
End Sub

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


For x% = 0 To 11
    If x <> 2 And x <> 6 And x <> 7 Then 'And x <> 8 Then
        If Not clpCode(x).ListChecker Then Exit Function
    End If
Next x%

For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

'------added on Sept 22, by Laura

For x% = 0 To 1
 If Len(dlpScheduled(x%).Text) > 0 Then
    If Not IsDate(dlpScheduled(x%).Text) Then
        MsgBox "Not a valid date"
        dlpScheduled(x%).Text = ""
        dlpScheduled(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

For x% = 0 To 1
 If Len(dlpStart(x%).Text) > 0 Then
    If Not IsDate(dlpStart(x%).Text) Then
        MsgBox "Not a valid date"
        dlpStart(x%).Text = ""
        dlpStart(x%).SetFocus
        Exit Function
    End If
 End If
Next x%


For x% = 0 To 1
 If Len(dlpRenewal(x%).Text) > 0 Then
    If Not IsDate(dlpRenewal(x%).Text) Then
        MsgBox "Not a valid date"
        dlpRenewal(x%).Text = ""
        dlpRenewal(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Len(dlpScheduled(0)) > 0 And Len(dlpScheduled(1)) > 0 Then
    If DaysBetween(dlpScheduled(0), dlpScheduled(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpScheduled(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If
If Len(dlpStart(0)) > 0 And Len(dlpStart(1)) > 0 Then
    If DaysBetween(dlpStart(0), dlpStart(1)) < 0 Then                                     '
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpStart(0).SetFocus                                             '
        Exit Function                                                       '
    End If
End If
If Len(dlpDateRange(0)) > 0 And Len(dlpDateRange(1)) > 0 Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                             '
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If
If Len(dlpRenewal(0)) > 0 And Len(dlpRenewal(1)) > 0 Then
    If DaysBetween(dlpRenewal(0), dlpRenewal(1)) < 0 Then                                 '
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpRenewal(0).SetFocus                                           '
        Exit Function                                                       '
    End If
End If

Call GetDisData

 '---------------------
If Not elpEEID.ListChecker Then Exit Function
If Not clpCode(2).CheckList Then Exit Function
If Not clpCode(12).CheckList Then Exit Function
If Not clpJOB.CheckList Then Exit Function

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

'Ticket #25306 - Jerry said this for everyone
'If glbWFC Then 'Ticket #25221 Franks 03/19/2014
    If optHideName.Value Then
        If comGroup(0).Text = "Employee Name" Or comGroup(0).Text = "Shift" Or comGroup(0).Text = lStr("Course Code") Then
            MsgBox "This report cannot sort by " & comGroup(0).Text
            Exit Function
        End If
    End If
'End If

CriCheck = True
End Function

Private Sub elpReptau_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMREDSEM"

Screen.MousePointer = HOURGLASS

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Region")

'Release 8.1 - Ticket #28040
comGroup(0).AddItem lStr("Rept. Authority 1")

'Release 8.0 - Ticket #22682
comGroup(0).AddItem lStr("Course Code")

comGroup(0).AddItem "(none)"
comGroup(1).AddItem "Employee Name"
comGroup(1).AddItem "Course Type"
comGroup(1).AddItem lStr("Course Code") 'Ticket #25148 Franks 03/04/2014
comGroup(1).AddItem "(none)"
comGroup(2).AddItem "Course Type"
comGroup(2).AddItem lStr("Course Code") 'Ticket #25148 Franks 03/04/2014
comGroup(2).AddItem "(none)"
comGroup(3).AddItem "Completed From/To"
comGroup(0).ListIndex = 0
comGroup(1).ListIndex = 0
comGroup(2).ListIndex = 0
comGroup(3).ListIndex = 0
comGroup(3).Enabled = False

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

lblKey.Caption = lStr(lblKey.Caption)
lblBCode(0).Caption = lStr(lblBCode(0).Caption)
lblBCode(1).Caption = lStr(lblBCode(1).Caption)
lblBCode(2).Caption = lStr(lblBCode(2).Caption)
lblBCode(3).Caption = lStr(lblBCode(3).Caption)
lblScheduled.Caption = lStr(lblScheduled.Caption)
lblStart.Caption = lStr(lblStart.Caption)
lblRenew.Caption = lStr(lblRenew.Caption)
lblTitle(22).Caption = lStr(lblTitle(22).Caption)
lblTitle(23).Caption = lStr(lblTitle(23).Caption)
lblTitle(24).Caption = lStr(lblTitle(24).Caption)
lblTitle(25).Caption = lStr(lblTitle(25).Caption)
lblRenewal.Caption = lStr(lblRenewal.Caption)
lblCourseName.Caption = lStr(lblCourseName.Caption)
lblCEU.Caption = lStr(lblCEU.Caption)

Call setRptCaption(Me)
Call setCaption(lblRepAuth)    'Ticket #27930 - Reporting Authority 1 from Current Position and not Attendance Supervisor

If glbLinamar Then clpCode(9).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(9).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

'Ticket #20447 - Jerry said to open up for everyone
'Frank added Apr 13, 2007 Ticket#12859
'If glbCompSerial = "S/N - 2188W" Then
    lblCEU.Visible = True
    clpCode(8).Visible = True
'End If

'Ticket #25306 - Jerry said this for everyone
'If Not glbWFC Then 'Ticket #25221 Franks 03/18/2014
'    optHideName.Visible = False
'End If
'If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'test only
'    optHideName.Enabled = False
'End If

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

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJOB.TransDiv = glbWFCUserSecList
    Call WFCScreenSetup 'Ticket #24708 Franks 04/08/2015
End If

''Hemu - 05/29/2003 Begin - Ticket # 4204
'If glbCompSerial = "S/N - 2161W" Then
'    clpCode(2).Width = 1400
''    clpCode(6).Width = 1400
''    clpCode(7).Width = 1400
''    clpCode(8).Width = 1400
'    clpCode(12).Width = 1400
'    clpCode(13).Width = 1400
'    clpCode(14).Width = 1400
'    clpCode(15).Width = 1400
'
'    clpCode(2).TextBoxWidth = 1200
'    clpCode(6).TextBoxWidth = 1200
'    clpCode(7).TextBoxWidth = 1200
'    clpCode(8).TextBoxWidth = 1200
'    clpCode(12).TextBoxWidth = 1200
'    clpCode(13).TextBoxWidth = 1200
'    clpCode(14).TextBoxWidth = 1200
'    clpCode(15).TextBoxWidth = 1200
'
'    clpCode(2).MaxLength = 8
'    clpCode(6).MaxLength = 8
'    clpCode(7).MaxLength = 8
'    clpCode(8).MaxLength = 8
'    clpCode(12).MaxLength = 8
'    clpCode(13).MaxLength = 8
'    clpCode(14).MaxLength = 8
'    clpCode(15).MaxLength = 8
'Else
'    clpCode(2).Width = 1200
''    clpCode(6).Width = 1200
''    clpCode(7).Width = 1200
''    clpCode(8).Width = 1200
'    clpCode(12).Width = 1200
''    clpCode(13).Width = 1200
''    clpCode(14).Width = 1200
''    clpCode(15).Width = 1200
'
'    clpCode(2).MaxLength = 4
''    clpCode(6).MaxLength = 4
''    clpCode(7).MaxLength = 4
''    clpCode(8).MaxLength = 4
'    clpCode(12).MaxLength = 4
'    clpCode(13).MaxLength = 4
'    clpCode(14).MaxLength = 4
'    clpCode(15).MaxLength = 4
'
'    clpCode(2).ShowDescription = False
'    clpCode(6).ShowDescription = False
'    clpCode(7).ShowDescription = False
'    clpCode(8).ShowDescription = False
'    clpCode(12).ShowDescription = False
'    clpCode(13).ShowDescription = False
'    clpCode(14).ShowDescription = False
'    clpCode(15).ShowDescription = False
'End If
'Hemu - 05/29/2003 End

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
    panWindow.Height = Me.ScaleHeight - (panControls.Height + 200)
    panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
    If panWindow.Height >= 12000 Then   '+ 230 Then
        scrControl.Value = 0
        panDetails.Top = 120
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
End Sub

Private Sub optCEU_Click(Value As Integer)
    'Ticket #22682 - Release 8.0
    If optSum.Value = True And chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked Then
        frCourseNameDesc.Visible = True
    Else
        frCourseNameDesc.Visible = False
    End If
    
    ''Ticket #25306 - Jerry said this for everyone
    'If glbWFC Then 'Ticket #25221 Franks 03/18/2014
        If chkNEKeyword.Enabled = False Then
            Call WFCHideNameFiledsNotUsed(True)
        End If
    'End If
    
    'Ticket #22682 - Release 8.0
    If optCEU.Value = True Then
        chkCEUCourseDescTrainer.Visible = True
    Else
        chkCEUCourseDescTrainer.Visible = False
    End If
    
    frHideShowCourseName.Visible = False 'Ticket #24708 Franks 04/08/2015
End Sub

Private Sub optCost_Click()
    If optCost.Value = True Then
      optDetail.Value = False
    ElseIf optCost.Value = False Then
      optDetail.Value = True
    End If

    'Ticket #22682 - Release 8.0
    If optSum.Value = True And chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked Then
        frCourseNameDesc.Visible = True
    Else
        frCourseNameDesc.Visible = False
    End If

    'Ticket #25306 - Jerry said this for everyone
    'If glbWFC Then 'Ticket #25221 Franks 03/18/2014
        If chkNEKeyword.Enabled = False Then
            Call WFCHideNameFiledsNotUsed(True)
        End If
    'End If
    
    'Ticket #22682 - Release 8.0
    If optCEU.Value = True Then
        chkCEUCourseDescTrainer.Visible = True
    Else
        chkCEUCourseDescTrainer.Visible = False
    End If
    
    frHideShowCourseName.Visible = False 'Ticket #24708 Franks 04/08/2015
End Sub

Private Sub optCost_GotFocus()
Call SetPanHelp(Me.ActiveControl)

End Sub

Private Sub optDetail_Click()
'If optDetail.Value = True Then
'  optCost.Value = False
'ElseIf optDetail.Value = False Then
'  optCost.Value = True
'End If

    'Ticket #22682 - Release 8.0
    If optSum.Value = True And chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked Then
        frCourseNameDesc.Visible = True
    Else
        frCourseNameDesc.Visible = False
    End If
    
    'Ticket #25306 - Jerry said this for everyone
    'If glbWFC Then 'Ticket #25221 Franks 03/18/2014
        If chkNEKeyword.Enabled = False Then
            Call WFCHideNameFiledsNotUsed(True)
        End If
    'End If
    
    'Ticket #22682 - Release 8.0
    If optCEU.Value = True Then
        chkCEUCourseDescTrainer.Visible = True
    Else
        chkCEUCourseDescTrainer.Visible = False
    End If
    
    frHideShowCourseName.Visible = False 'Ticket #24708 Franks 04/08/2015
End Sub

Private Sub optDetail_GotFocus()
Call SetPanHelp(Me.ActiveControl)      'laura jan 06, 1998
End Sub

Private Sub optHideName_Click(Value As Integer)
    Call WFCHideNameFiledsNotUsed(False)

    'Ticket #22682 - Release 8.0
    If optCEU.Value = True Then
        chkCEUCourseDescTrainer.Visible = True
        'frHideShowCourseName.Visible = False 'Ticket #24708 Franks 04/08/2015
    Else
        chkCEUCourseDescTrainer.Visible = False
        'frHideShowCourseName.Visible = True 'Ticket #24708 Franks 04/08/2015
    End If
    frHideShowCourseName.Visible = True 'Ticket #24708 Franks 04/08/2015
    frCourseNameDesc.Visible = False 'Ticket #24708 Franks 04/08/2015
End Sub

Private Sub optSum_Click()
    'Ticket #22682 - Release 8.0
    If optSum.Value = True And chkNEKeyword.Value = vbUnchecked And chkNECourseType.Value = vbUnchecked And chkNECourseCode.Value = vbUnchecked And chkNEResults.Value = vbUnchecked And chkNEConduct.Value = vbUnchecked Then
        frCourseNameDesc.Visible = True
    Else
        frCourseNameDesc.Visible = False
    End If
    
    'Ticket #25306 - Jerry said this for everyone
    'If glbWFC Then 'Ticket #25221 Franks 03/18/2014
        If chkNEKeyword.Enabled = False Then
            Call WFCHideNameFiledsNotUsed(True)
        End If
    'End If
    
    'Ticket #22682 - Release 8.0
    If optCEU.Value = True Then
        chkCEUCourseDescTrainer.Visible = True
    Else
        chkCEUCourseDescTrainer.Visible = False
    End If
    
    frHideShowCourseName.Visible = False 'Ticket #24708 Franks 04/08/2015
End Sub

Private Sub optSum_GotFocus()
Call SetPanHelp(Me.ActiveControl)      'laura jan 06, 1998
End Sub

Private Sub scrControl_Change()
    panDetails.Top = 120 - scrControl.Value
End Sub

Private Sub txtCourseName_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtKey_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function getCourseDesc(xTBName, xTBCode)
Dim rsTemp As New ADODB.Recordset
If xTBName = "JOB" Then
    rsTemp.Open "SELECT JB_DESCR AS FIND_DESC FROM HRJOB WHERE JB_CODE='" & xTBCode & "'", gdbAdoIhr001, adOpenForwardOnly
Else
    rsTemp.Open "SELECT TB_DESC AS FIND_DESC FROM HRTABL WHERE TB_NAME='" & xTBName & "' AND TB_KEY='" & xTBCode & "'", gdbAdoIhr001, adOpenForwardOnly
End If

getCourseDesc = ""
If Not rsTemp.EOF Then
    getCourseDesc = rsTemp("FIND_DESC") & ""
End If
End Function

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

Private Sub Cri_Trainer()
Dim EECri As String, OneSet%, x%
Dim strCa$

If Len(txtTrainerName.Text) < 1 Then Exit Sub

strCa$ = "HREDSEM.ES_TRAINNER"

If Len(txtTrainerName.Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} >= '" & Replace(txtTrainerName, "'", "''") & "') "
End If


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_CompanyName()
Dim EECri As String, OneSet%, x%
Dim strCa$

If Len(txtCompanyName.Text) < 1 Then Exit Sub

strCa$ = "HREDSEM.ES_COMPANYNAME"

If Len(txtCompanyName.Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} >= '" & Replace(txtCompanyName, "'", "''") & "') "
End If


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_Coord()

    Dim EECri As String, OneSet%, x%
    Dim strCa$
    
    If Len(clpCode(6).Text) < 1 Then Exit Sub
    
    strCa$ = "HREDSEM.ES_COORDINATED"

    If Len(clpCode(6).Text) > 0 Then
      EECri = EECri & "({" & strCa$ & "} IN ['" & Replace(clpCode(6).Text, ",", "','") & "'])"
      WrkTabSQL = WrkTabSQL & " AND " & EECri
    End If
    

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_Method()

    Dim EECri As String, OneSet%, x%
    Dim strCa$
    
    If Len(clpCode(7).Text) < 1 Then Exit Sub
    
    strCa$ = "HREDSEM.ES_METHODUSED"

    If Len(clpCode(7).Text) > 0 Then
      EECri = EECri & "({" & strCa$ & "} IN ['" & Replace(clpCode(7).Text, ",", "','") & "'])"
      WrkTabSQL = WrkTabSQL & " AND " & EECri
    End If
    

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub setRptLabelEdSem()
    vbxCrystal.Formulas(40) = "lblCourseType='" & lStr("Course Type") & "'"
    vbxCrystal.Formulas(41) = "lblCourseCode='" & lStr("Course Code") & "'"
    vbxCrystal.Formulas(42) = "lblCondBy='" & lStr("Conducted By") & "'"
    vbxCrystal.Formulas(43) = "lblResults='" & lStr("Results") & "'"
    vbxCrystal.Formulas(44) = "lblKeyword='" & lStr("Keyword") & "'"
End Sub

Private Sub setRptLabelEdSemMore()
    vbxCrystal.Formulas(45) = "lblEmployee='" & lStr("Employee $") & "'"
    vbxCrystal.Formulas(46) = "lblEmployer='" & lStr("Employer $") & "'"
    vbxCrystal.Formulas(47) = "lblOther='" & lStr("Other Expenses $") & "'"
    vbxCrystal.Formulas(48) = "lblAccom='" & lStr("Accomodation $") & "'"
    vbxCrystal.Formulas(49) = "lblLearn='" & lStr("Learning Material $") & "'"
End Sub

Private Sub setRptLabelEdSem_More()
    vbxCrystal.Formulas(40) = "lblCourseType='" & lStr("Course Type") & "'"
    vbxCrystal.Formulas(41) = "lblCourseCode='" & lStr("Course Code") & "'"
    vbxCrystal.Formulas(42) = "lblCondBy='" & lStr("Conducted By") & "'"
    vbxCrystal.Formulas(43) = "lblResults='" & lStr("Results") & "'"
    vbxCrystal.Formulas(44) = "lblKeyword='" & lStr("Keyword") & "'"
    vbxCrystal.Formulas(43) = "lblCompName='" & lStr("Company Name") & "'"
    vbxCrystal.Formulas(44) = "lblTrainer='" & lStr("Trainer") & "'"
End Sub

Private Sub WFCHideNameFiledsNotUsed(xFlag As Boolean) 'Ticket #25221 Franks 03/18/2014
    If Not xFlag Then
        chkNEKeyword.Value = False
        chkNECourseType.Value = False
        chkNECourseCode.Value = False
        chkNEConduct.Value = False
        chkNoStartDate.Value = False
        chkNoDateComplete.Value = False
        chkNEResults.Value = False
    End If
    chkNEKeyword.Enabled = xFlag
    chkNECourseType.Enabled = xFlag
    chkNECourseCode.Enabled = xFlag
    chkNEConduct.Enabled = xFlag
    chkNoStartDate.Enabled = xFlag
    chkNoDateComplete.Enabled = xFlag
    chkNEResults.Enabled = xFlag
    lblShift.Enabled = xFlag
    txtShift.Enabled = xFlag
    Label1.Enabled = xFlag
    clpProv.Enabled = xFlag
    Label2.Enabled = xFlag
    clpProvEmp.Enabled = xFlag
    lblJOB.Enabled = xFlag
    clpJOB.Enabled = xFlag
    lblGrp(1).Enabled = xFlag
    comGroup(1).Enabled = xFlag
    lblGrp(2).Enabled = xFlag
    comGroup(2).Enabled = xFlag
    lblGrp(3).Enabled = xFlag
End Sub

Private Sub WFCHideNameRpt() 'Ticket #25221 Franks 03/19/2014
Dim rsEdSem As New ADODB.Recordset
Dim rsEdWrk As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim SQLQ As String
Dim xGroup, xGroupField, xGrpVal, xGrpDesc
Dim I As Long
Dim xtot As Long
Dim xCrsCode, xCType, xDate
Dim xEmpnbr, xPreEmpNo
Dim xCurGrp, xPreGrp
Dim xGrandEmpCnt, xGrpEmpCnt
Dim xSeleSQL

    SQLQ = "DELETE FROM HREDWRK WHERE ES_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ

    WrkTabSQL = Replace(Replace(WrkTabSQL, "}", ""), "{", "")
    WrkTabSQL = Replace(Replace(WrkTabSQL, "[", "("), "]", ")")
    SQLQ = WrkTabSQL
    If Len(DateSelCri) > 0 Then
        SQLQ = SQLQ & " AND " & DateSelCri
    End If
    'CAN'T BE Shift and Course and Name
    xGroup = comGroup(0).Text
    xGroupField = ""
    If xGroup = lStr("Division") Then xGroupField = "ED_DIV"
    If xGroup = lStr("Department") Then xGroupField = "ED_DEPTNO"
    If xGroup = lStr("Location") Then xGroupField = "ED_LOC"
    If xGroup = lStr("Section") Then xGroupField = "ED_SECTION"
    If xGroup = lStr("Region") Then xGroupField = "ED_REGION"
    'If xGroup = "(none)" Then nothing
    
    xSeleSQL = SQLQ
    SQLQ = "SELECT HREDSEM.*,ED_DIV, ED_DEPTNO,ED_LOC,ED_SECTION,ED_REGION, ED_SURNAME, ED_FNAME FROM HREDSEM LEFT JOIN HREMP ON HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR WHERE " & xSeleSQL & " "
    SQLQ = SQLQ & "AND NOT (ES_CRSCODE IS NULL) "
    If Len(xGroupField) > 0 Then
        SQLQ = SQLQ & "ORDER BY " & xGroupField & ",ES_CRSCODE, ES_EMPNBR "
    Else
        SQLQ = SQLQ & "ORDER BY ES_CRSCODE, ES_EMPNBR "
    End If
    
    xGrandEmpCnt = 0
    If rsEdSem.State <> 0 Then rsEdSem.Close
    rsEdSem.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsEdSem.EOF Then
        MsgBox "There is no record found."
        Exit Sub
    Else
        I = 0
        xtot = rsEdSem.RecordCount
        SQLQ = "SELECT COUNT(DISTINCT ES_EMPNBR) AS EMPTOT FROM HREDSEM LEFT JOIN HREMP ON HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR WHERE " & xSeleSQL & " "
        SQLQ = SQLQ & "AND NOT (ES_CRSCODE IS NULL) "
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTmp.EOF Then
            xGrandEmpCnt = rsTmp("EMPTOT")
        End If
        rsTmp.Close
    End If
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 1
    xPreEmpNo = 0
    xPreGrp = "*"
    Do While Not rsEdSem.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xtot) * 100
        I = I + 1
        DoEvents
        If Len(xGroupField) > 0 Then
            If IsNull(rsEdSem(xGroupField)) Then
                xCurGrp = ""
            Else
                xCurGrp = rsEdSem(xGroupField)
            End If
            If Not xCurGrp = xPreGrp Then
                If Not xPreGrp = "*" Then
                    'getTotEmp for this Grp
                    SQLQ = "SELECT COUNT(DISTINCT ES_EMPNBR) AS EMPTOT FROM HREDSEM LEFT JOIN HREMP ON HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR WHERE " & xSeleSQL & " "
                    SQLQ = SQLQ & "AND NOT (ES_CRSCODE IS NULL) "
                    If Len(xPreGrp) = 0 Then
                        SQLQ = SQLQ & "AND " & xGroupField & " IS NULL "
                    Else
                        SQLQ = SQLQ & "AND " & xGroupField & " = '" & xPreGrp & "' "
                    End If
                    If rsTmp.State <> 0 Then rsTmp.Close
                    rsTmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                    xGrpEmpCnt = 0
                    If Not rsTmp.EOF Then
                        xGrpEmpCnt = rsTmp("EMPTOT")
                    End If
                    rsTmp.Close
                    
                    SQLQ = "UPDATE HREDWRK SET ES_NUM2 = " & xGrpEmpCnt & " WHERE ES_WRKEMP = '" & glbUserID & "' "
                    SQLQ = SQLQ & "AND ES_RPT_GRP = '" & xPreGrp & "' "
                    gdbAdoIhr001.Execute SQLQ
                End If
                xPreGrp = xCurGrp
            End If
        End If
        
        SQLQ = "SELECT * FROM HREDWRK WHERE ES_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND ES_CRSCODE = '" & rsEdSem("ES_CRSCODE") & "' "
        If IsNull(rsEdSem("ES_DATCOMP")) Then
            SQLQ = SQLQ & "AND ES_DATCOMP IS NULL "
        Else
            SQLQ = SQLQ & "AND ES_DATCOMP = " & Date_SQL(rsEdSem("ES_DATCOMP")) & " "
        End If
        If Len(xGroupField) > 0 Then
            If IsNull(rsEdSem(xGroupField)) Then
                SQLQ = SQLQ & "AND ES_RPT_GRP IS NULL "
            Else
                SQLQ = SQLQ & "AND ES_RPT_GRP  = '" & rsEdSem(xGroupField) & "' "
            End If
        End If
        xEmpnbr = rsEdSem("ES_EMPNBR")
        If rsEdWrk.State <> 0 Then rsEdWrk.Close
        rsEdWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEdWrk.EOF Then
            rsEdWrk.AddNew
            rsEdWrk("ES_CRSCODE") = rsEdSem("ES_CRSCODE")
            rsEdWrk("ES_CRSDESC") = Left(GetTABLCodePub("ESCD", rsEdSem("ES_CRSCODE")), 30)
            rsEdWrk("ES_CTYPE") = rsEdSem("ES_CTYPE")
            rsEdWrk("ES_DATCOMP") = rsEdSem("ES_DATCOMP")
            rsEdWrk("ES_HOURS") = 0
            rsEdWrk("ES_TBEMP") = 0
            rsEdWrk("ES_TBCO") = 0
            rsEdWrk("ES_OTHER") = 0
            rsEdWrk("ES_ACCOM") = 0
            xPreEmpNo = xEmpnbr
            rsEdWrk("ES_NUM1") = 1 '0 'count # of employee
            rsEdWrk("ES_NUM2") = 0
            rsEdWrk("ES_NUM3") = 0
            rsEdWrk("ES_WRKEMP") = glbUserID
            If Len(xGroupField) > 0 Then
                If Not IsNull(rsEdSem(xGroupField)) Then
                    rsEdWrk("ES_RPT_GRP") = Left(rsEdSem(xGroupField), 20)
                    rsEdWrk("ES_RPT_GRPDESC") = getHideNameRptGrpDesc(xGroupField, rsEdSem(xGroupField))
                End If
            End If
            
        Else
        '    MsgBox ""
        End If
        If Not IsNull(rsEdSem("ES_HOURS")) Then rsEdWrk("ES_HOURS") = rsEdWrk("ES_HOURS") + rsEdSem("ES_HOURS")
        If Not IsNull(rsEdSem("ES_TBEMP")) Then rsEdWrk("ES_TBEMP") = rsEdWrk("ES_TBEMP") + rsEdSem("ES_TBEMP")
        If Not IsNull(rsEdSem("ES_TBCO")) Then rsEdWrk("ES_TBCO") = rsEdWrk("ES_TBCO") + rsEdSem("ES_TBCO")
        If Not IsNull(rsEdSem("ES_OTHER")) Then rsEdWrk("ES_OTHER") = rsEdWrk("ES_OTHER") + rsEdSem("ES_OTHER")
        'others = ES_OTHER + ES_ACCOM + ES_LEARNING
        If Not IsNull(rsEdSem("ES_ACCOM")) Then rsEdWrk("ES_OTHER") = rsEdWrk("ES_OTHER") + rsEdSem("ES_ACCOM")
        If Not IsNull(rsEdSem("ES_LEARNING")) Then rsEdWrk("ES_OTHER") = rsEdWrk("ES_OTHER") + rsEdSem("ES_LEARNING")
        'If Not IsNull(rsEdSem("ES_ACCOM")) Then rsEdWrk("ES_ACCOM") = rsEdWrk("ES_ACCOM") + rsEdSem("ES_ACCOM")
        If Not xEmpnbr = xPreEmpNo Then
            rsEdWrk("ES_NUM1") = rsEdWrk("ES_NUM1") + 1
            xPreEmpNo = xEmpnbr
        End If
        rsEdWrk.Update
        rsEdWrk.Close
        
        rsEdSem.MoveNext
    Loop

    If Not xPreGrp = "*" Then
        'getTotEmp for this Grp
        SQLQ = "SELECT COUNT(DISTINCT ES_EMPNBR) AS EMPTOT FROM HREDSEM LEFT JOIN HREMP ON HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR WHERE " & xSeleSQL & " "
        SQLQ = SQLQ & "AND NOT (ES_CRSCODE IS NULL) "
        If Len(xPreGrp) = 0 Then
            SQLQ = SQLQ & "AND " & xGroupField & " IS NULL "
        Else
            SQLQ = SQLQ & "AND " & xGroupField & " = '" & xPreGrp & "' "
        End If
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        xGrpEmpCnt = 0
        If Not rsTmp.EOF Then
            xGrpEmpCnt = rsTmp("EMPTOT")
        End If
        rsTmp.Close

        SQLQ = "UPDATE HREDWRK SET ES_NUM2 = " & xGrpEmpCnt & " WHERE ES_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND ES_RPT_GRP = '" & xPreGrp & "' "
        gdbAdoIhr001.Execute SQLQ
    End If
    'Grand Total
    SQLQ = "UPDATE HREDWRK SET ES_NUM3 = " & xGrandEmpCnt & " WHERE ES_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ

    rsEdSem.Close
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub

Private Function getHideNameRptGrpDesc(xFldName, xVal)
Dim retval
    retval = ""
    If xFldName = "ED_DIV" Then
        retval = getDivDescPub(xVal)
    End If
    If xFldName = "ED_DEPTNO" Then
        retval = getDeptDescPub(xVal)
    End If
    If xFldName = "ED_LOC" Then
        retval = GetTABLCodePub("EDLC", xVal)
    End If
    If xFldName = "ED_SECTION" Then
        retval = GetTABLCodePub("EDSE", xVal)
    End If
    If xFldName = "ED_REGION" Then
        retval = GetTABLCodePub("EDRG", xVal)
    End If
    getHideNameRptGrpDesc = retval

End Function

Private Sub WFCScreenSetup() 'Ticket #24708 Franks 04/08/2015
    frHideShowCourseName.Left = 5880
    frHideShowCourseName.Top = frCourseNameDesc.Top
    frHideShowCourseName.BorderStyle = 0
End Sub
