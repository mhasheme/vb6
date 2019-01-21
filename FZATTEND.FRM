VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRAttend 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   10725
   ClientLeft      =   525
   ClientTop       =   1200
   ClientWidth     =   14010
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
   ScaleHeight     =   10725
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   90
      Top             =   9120
      Width           =   12255
   End
   Begin VB.VScrollBar scrControl 
      Height          =   8715
      LargeChange     =   315
      Left            =   12240
      Max             =   100
      SmallChange     =   315
      TabIndex        =   89
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12135
      Begin VB.CheckBox chkUseEmpMst 
         Caption         =   "Use Employee Master"
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
         Left            =   4150
         TabIndex        =   44
         Tag             =   "Use data from Employee Master"
         Top             =   7245
         Visible         =   0   'False
         Width           =   1880
      End
      Begin VB.CheckBox chkInclTerm 
         Caption         =   "Include Terminated Employees"
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
         Left            =   6240
         TabIndex        =   100
         Top             =   7360
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7935
         TabIndex        =   35
         Top             =   6165
         Width           =   870
      End
      Begin VB.CheckBox chkUseHireDate 
         Caption         =   "Use Original Date of Hire"
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
         Left            =   6240
         TabIndex        =   40
         Top             =   7125
         Width           =   2175
      End
      Begin VB.TextBox TxtCases 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   7935
         TabIndex        =   94
         Top             =   5813
         Width           =   2052
      End
      Begin VB.ComboBox cmbAttendGroup 
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
         Left            =   2115
         TabIndex        =   93
         Top             =   5800
         Width           =   2652
      End
      Begin VB.TextBox txtIncentive 
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
         Left            =   3795
         MaxLength       =   1
         TabIndex        =   19
         Tag             =   "Incentive - Y or N"
         Top             =   4494
         Width           =   435
      End
      Begin VB.TextBox txtHoursLess 
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
         Left            =   7935
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "10-Enter Hours"
         Top             =   4505
         Width           =   870
      End
      Begin VB.TextBox txtChargeCode 
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
         Left            =   2115
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "00-Specific Department Desired"
         Top             =   3222
         Width           =   855
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
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "00-Shift individual is on"
         Top             =   4176
         Width           =   780
      End
      Begin VB.TextBox txtIncident 
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
         Left            =   2115
         MaxLength       =   1
         TabIndex        =   18
         Tag             =   "Incident - Y or N"
         Top             =   4494
         Width           =   435
      End
      Begin VB.TextBox txtHours 
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
         Left            =   7935
         MaxLength       =   7
         TabIndex        =   26
         Tag             =   "10-Enter Hours"
         Top             =   4188
         Width           =   870
      End
      Begin VB.TextBox txtSeniority 
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
         Left            =   7935
         MaxLength       =   1
         TabIndex        =   28
         Tag             =   "30-seniority Flag - Y or N"
         Top             =   4822
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Tag             =   "First Level of grouping records"
         Top             =   7200
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Tag             =   "Second level of grouping records"
         Top             =   7525
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "Third level of grouping records"
         Top             =   7850
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
         Index           =   3
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "Final sorting of records - no totals"
         Top             =   8175
         Width           =   2325
      End
      Begin VB.CheckBox chkComments 
         Caption         =   "Show Comments"
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
         Left            =   3405
         TabIndex        =   37
         Top             =   6600
         Width           =   2055
      End
      Begin VB.CheckBox chkCostHis 
         Caption         =   "Cost At Historical Rates"
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
         Left            =   9120
         TabIndex        =   41
         Top             =   6600
         Value           =   1  'Checked
         Width           =   2055
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
         Height          =   225
         Left            =   6240
         TabIndex        =   38
         Top             =   6600
         Width           =   2055
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
         Left            =   7935
         TabIndex        =   23
         Tag             =   "Choose Upload flag."
         Top             =   3207
         Width           =   975
      End
      Begin VB.CheckBox chkInclHIS 
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
         Height          =   285
         Left            =   9120
         TabIndex        =   42
         Top             =   6832
         Width           =   2355
      End
      Begin VB.CheckBox chkSup 
         Caption         =   "Use current Reporting Authority"
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
         Left            =   6240
         TabIndex        =   39
         Top             =   6862
         Width           =   2655
      End
      Begin INFOHR_Controls.CodeLookup clpJOB 
         Height          =   285
         Left            =   7620
         TabIndex        =   24
         Tag             =   "00-Enter Position Code"
         Top             =   3555
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.EmployeeLookup elpSUP 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Tag             =   "00-Employee Number "
         Top             =   2586
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   10
         Left            =   1800
         TabIndex        =   21
         Tag             =   "00-Enter Administered By Code"
         Top             =   5130
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   20
         Tag             =   "00-Enter Region Code"
         Top             =   4812
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   8
         Left            =   7620
         TabIndex        =   29
         Tag             =   "00-Enter Section Code"
         Top             =   5145
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   7
         Left            =   7620
         TabIndex        =   25
         Tag             =   "00-Enter Position Group"
         Top             =   3870
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   5
         Tag             =   "00-Enter Status Code"
         Top             =   1632
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin MSMask.MaskEdBox medDOW 
         Height          =   285
         Index           =   0
         Left            =   2115
         TabIndex        =   14
         Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
         Top             =   3858
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   0
         Left            =   1845
         TabIndex        =   31
         Tag             =   "Detailed Report"
         Top             =   6180
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
         Value           =   -1  'True
      End
      Begin MSMask.MaskEdBox medDOW 
         Height          =   285
         Index           =   1
         Left            =   3075
         TabIndex        =   15
         Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
         Top             =   3858
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   1
         Left            =   3300
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "Summary Report"
         Top             =   6180
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Summary"
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
      Begin Threed.SSFrame fraDOW 
         Height          =   1905
         Left            =   9720
         TabIndex        =   48
         Top             =   300
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   3360
         _StockProps     =   14
         Caption         =   "Days of the Week"
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
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Monday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   55
            Top             =   525
            Width           =   960
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Tuesday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   54
            Top             =   735
            Width           =   1020
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "4 - Wednesday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   330
            TabIndex        =   53
            Top             =   945
            Width           =   1290
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "5 - Thursday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   330
            TabIndex        =   52
            Top             =   1155
            Width           =   1080
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "6 - Friday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   330
            TabIndex        =   51
            Top             =   1365
            Width           =   810
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "7 - Saturday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   330
            TabIndex        =   50
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblDOW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Sunday"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   49
            Top             =   320
            Width           =   930
         End
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Tag             =   "00-Specific Department Desired"
         Top             =   678
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
         Left            =   1800
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   360
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
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   996
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
         Top             =   1950
         Width           =   7515
         _ExtentX        =   13256
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
         Left            =   3450
         TabIndex        =   13
         Tag             =   "40-Date upto and including this date forward"
         Top             =   3540
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
         TabIndex        =   12
         Tag             =   "40-Date from and including this date forward"
         Top             =   3540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.CodeLookup clpChrgCode 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Tag             =   "00-Specific Department Desired"
         Top             =   3222
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   17
         Tag             =   "00-Fund"
         Top             =   4176
         Visible         =   0   'False
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   56
         Tag             =   "01-Counselling Type- Code"
         Top             =   8640
         Visible         =   0   'False
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "CETY"
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Tag             =   "00-Enter Union Code"
         Top             =   1314
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Tag             =   "ADRE-Attendance Reason"
         Top             =   2904
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         DataField       =   "AD_MACHINE_NUM"
         Height          =   285
         Index           =   5
         Left            =   7620
         TabIndex        =   30
         Tag             =   "00-Machine #"
         Top             =   5460
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   12
      End
      Begin INFOHR_Controls.CodeLookup clpcode 
         DataField       =   "AD_PROJECT_CODE"
         Height          =   285
         Index           =   11
         Left            =   1800
         TabIndex        =   22
         Tag             =   "00-Account Code"
         Top             =   5460
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   11
         LookupType      =   11
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Tag             =   "10-Enter Employee Number"
         Top             =   2268
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "Summary Report"
         Top             =   6180
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   For Payroll"
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
      Begin VB.TextBox txtWeek 
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
         Left            =   10215
         TabIndex        =   36
         Top             =   6165
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   3
         Left            =   1845
         TabIndex        =   34
         TabStop         =   0   'False
         Tag             =   "Summary Report"
         Top             =   6585
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Employee Totals Only"
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
      Begin VB.Label lblYear 
         Caption         =   "Year"
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
         Left            =   6240
         TabIndex        =   99
         Top             =   6210
         Width           =   1395
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   9900
         Picture         =   "FZATTEND.frx":0000
         Top             =   6180
         Visible         =   0   'False
         Width           =   240
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
         TabIndex        =   97
         Top             =   2313
         Width           =   1290
      End
      Begin VB.Label lblCodes 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Codes"
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
         TabIndex        =   96
         Top             =   5830
         Width           =   1215
      End
      Begin VB.Label LblCases 
         Caption         =   "Case Files"
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
         Left            =   6240
         TabIndex        =   95
         Top             =   5830
         Width           =   1335
      End
      Begin VB.Label lblIncentive 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive"
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
         Left            =   3000
         TabIndex        =   92
         Top             =   4539
         Width           =   660
      End
      Begin VB.Label lblHoursLess 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours ' <= '"
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
         Left            =   6240
         TabIndex        =   91
         Top             =   4545
         Width           =   795
      End
      Begin VB.Label lblAttendCrit 
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
         TabIndex        =   88
         Top             =   0
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
         Left            =   120
         TabIndex        =   87
         Top             =   405
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
         TabIndex        =   86
         Top             =   723
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
         TabIndex        =   85
         Top             =   1359
         Width           =   420
      End
      Begin VB.Label lblStatus 
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
         Left            =   120
         TabIndex        =   84
         Top             =   1677
         Width           =   450
      End
      Begin VB.Label lblChargeCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Code"
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
         TabIndex        =   83
         Top             =   3267
         Width           =   930
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
         TabIndex        =   82
         Top             =   3585
         Width           =   1095
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Week"
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
         TabIndex        =   81
         Top             =   3903
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
         Left            =   120
         TabIndex        =   80
         Top             =   4221
         Width           =   315
      End
      Begin VB.Label lblIncident 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident"
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
         TabIndex        =   79
         Top             =   4539
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
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
         TabIndex        =   78
         Top             =   6210
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
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
         Left            =   2700
         TabIndex        =   77
         Top             =   3903
         Width           =   270
      End
      Begin VB.Label lblHours 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours ' >= '"
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
         Left            =   6240
         TabIndex        =   76
         Top             =   4230
         Width           =   795
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
         Left            =   120
         TabIndex        =   75
         Top             =   2631
         Width           =   750
      End
      Begin VB.Label lblSen1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seniority"
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
         Left            =   6240
         TabIndex        =   74
         Top             =   4860
         Width           =   600
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
         Left            =   120
         TabIndex        =   73
         Top             =   6885
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
         TabIndex        =   72
         Top             =   7260
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   180
         TabIndex        =   71
         Top             =   7585
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Index           =   2
         Left            =   180
         TabIndex        =   70
         Top             =   7910
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
         TabIndex        =   69
         Top             =   8235
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
         TabIndex        =   68
         Top             =   1041
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
         TabIndex        =   67
         Top             =   4857
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
         TabIndex        =   66
         Top             =   5175
         Width           =   1125
      End
      Begin VB.Label lblPosGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Group"
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
         Left            =   6240
         TabIndex        =   65
         Top             =   3915
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   6240
         TabIndex        =   64
         Top             =   5190
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
         TabIndex        =   63
         Top             =   1995
         Width           =   630
      End
      Begin VB.Label lblJOB 
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
         Left            =   6240
         TabIndex        =   62
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Flag"
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
         Left            =   6240
         TabIndex        =   61
         Top             =   3270
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disciplinary"
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
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   60
         Top             =   8655
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   59
         Top             =   2949
         Width           =   1320
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
         Left            =   6240
         TabIndex        =   58
         Top             =   5475
         Width           =   1215
      End
      Begin VB.Label lblAccount 
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
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   5505
         Width           =   1215
      End
      Begin VB.Label lblWeek 
         Caption         =   "Pay Period #"
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
         Left            =   8880
         TabIndex        =   98
         Top             =   6210
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8400
      Top             =   9840
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
      Left            =   8880
      Top             =   9840
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
Attribute VB_Name = "frmRAttend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'changed optVDesc(index) to  cmdVDesc(index) 'Jaddy 10/24/99
Option Explicit
Dim ReportSel, HisSQL
Dim fDis(10) As String
Dim flgUseLabelMst As String
', HisSQL1
'Dim SALCostHisSQL, ATTCostHisSQL
'Dim WithEvents CN001 As ADODB.Connection

Private Sub chkInclHIS_Click()
    If Not gSec_Inq_Attendance_History And chkInclHIS.Value = 1 Then
        MsgBox "You Do Not Have Authority For This Transaction"
        chkInclHIS.Value = 0
    End If
End Sub

Private Sub chkUseEmpMst_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpChrgCode_Change()
    txtChargeCode = clpChrgCode
End Sub

Private Sub clpCode_Change(Index As Integer)
    If Index = 3 Then
        txtShift = clpCode(3)
    End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
'on error Resume Next: lblCodeDesc(Index).Caption = clpCode(Index).Caption
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Attendance Report Criteria", Me) Then Exit Sub
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    If ReportSel <> "EnviroServRpt" Then
        Me.vbxCrystal.Destination = 1
        MDIMain.Timer1.Enabled = False
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
    End If
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
     Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X%, SQLQ
Dim strWHand As String, MyQuery As QueryDef
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
    If ReportSel <> "EnviroServRpt" Then
        Me.vbxCrystal.Destination = 0
        MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
    End If
    'added by Bryan Ticket#12094 access users need some time between reports
    If Not (glbOracle Or glbSQL) Then Pause (10)
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
'MsgBox Error$(Err)
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If MDIMain.panHelp(0).FloodType > 0 Then MDIMain.panHelp(0).FloodType = 0  'ADDED BY RAUBREY 4/10/97
Resume Next
Screen.MousePointer = DEFAULT

End Sub

Private Sub comGroup_Click(Index As Integer)
    If comGroup(0).Text = lStr("G/L") Then
        'Ticket #29984 - Option to use GL No from Employee Master instead of Attendance data
        chkUseEmpMst.Visible = True
    Else
        chkUseEmpMst.Visible = False
    End If
End Sub

'Private Sub cmdView_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub comGrpLoad()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

    If frmRAttend.Caption = "Costed Attendance Report" Then
        lblPosGroup.Visible = True
        clpCode(7).Visible = True
    End If
    
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(3).Clear
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("G/L")
    
    If frmRAttend.Caption = "Costed Attendance Report" Then 'Frank 4/20/2000
        Dim vPosGroup As String
        If Not glbSyndesis Then
            vPosGroup = "Position Group"
        Else
            vPosGroup = "Position Grade"
        End If
        comGroup(0).AddItem vPosGroup '"Position Group"
    End If
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Position"     'Hemu - Jerry asked to add this
    comGroup(0).AddItem lStr("AttSupervisor")
    
    If frmRAttend.Caption = "Attendance Report" Then 'Ticket #21654 Franks 03/22/2012
        comGroup(0).AddItem lStr("Administered By")
    End If
    
    'Ticket #26118 - Add Account Code to G1 and G2
    comGroup(0).AddItem lStr("Account Code")
    
    comGroup(0).AddItem "(none)"
    
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Attendance Code"
    'Ticket #26118 - Add Account Code to G1 and G2
    comGroup(1).AddItem lStr("Account Code")
    
    If frmRAttend.Caption = "Attendance Report" Or frmRAttend.Caption = "Attendance Report Including Historical Data" Then 'Ticket #21654 Franks 03/22/2012
        comGroup(1).AddItem lStr("Section")
    End If
    comGroup(1).AddItem "Position"     'Release 8.1
    comGroup(1).AddItem "(none)"
    
    comGroup(2).AddItem "Attendance Code"
    comGroup(2).AddItem "(none)"
    
    comGroup(3).AddItem "Attendance Date"
    comGroup(3).AddItem "(none)"
    
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    
    comGroup(3).Enabled = False
    
    ReportSel = "ATT"
    
    If Me.Caption = "Attendance Report" Then ReportSel = "ATT"
    If Me.Caption = "Attendance Report Including Historical Data" Then ReportSel = "HIS"
    If Me.Caption = "Costed Attendance Report" Then ReportSel = "COS"
    If Me.Caption = "Personal Day Report" Then ReportSel = "PerDayRpt"
    If Me.Caption = "Wellington Terrace Attendance Report" Then ReportSel = "EnviroServRpt"    'Ticket #22034
    
    If ReportSel = "COS" Then
        If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then
            chkCostHis.Visible = False
            chkCostHis.Value = 0
            chkInclHIS.Visible = False
            chkInclHIS.Value = 0
        Else
            chkCostHis.Visible = True
            chkInclHIS.Visible = True
            chkCostHis.Value = 1
        End If
        
        'Ticket #21631
        chkUseHireDate.Visible = False
        
        'Ticket #24655 - WDGPHU
        If glbCompSerial = "S/N - 2411W" Then
            optGrouping(2).Visible = True
            lblYear.Visible = True
            txtYear.Visible = True
            lblWeek.Visible = True
            txtWeek.Visible = True
            imgIcon.Visible = True
        End If
    Else
        chkCostHis.Visible = False
        chkCostHis.Value = 0
        chkInclHIS.Visible = False
        chkInclHIS.Value = 0
        
        'Ticket #21631
        chkUseHireDate.Visible = True
        
        'Ticket #24655 - WDGPHU
        If glbCompSerial = "S/N - 2411W" Then
            optGrouping(2).Visible = False
            lblYear.Visible = False
            txtYear.Visible = False
            lblWeek.Visible = False
            txtWeek.Visible = False
            imgIcon.Visible = False
        End If
    End If
    
    If glbCompSerial = "S/N - 2241W" Then
        If ReportSel = "PerDayRpt" Then 'Ticket #15133
            clpCode(2).Enabled = False
            clpCode(2).Text = "PERS" '"VAC,SICF"
            Call ScreenPerDayRpt
        Else
            clpCode(2).Enabled = True
            clpCode(2).Text = ""
        End If
    End If
    
    If ReportSel = "EnviroServRpt" Then   'Ticket #22034
        lblTitle(8).Visible = False
        optGrouping(0).Visible = False
        optGrouping(1).Visible = False
        optGrouping(3).Visible = False
        chkComments.Visible = False
        chkSup.Visible = False
        chkUseHireDate.Value = vbUnchecked
        chkUseHireDate.Visible = False
        lblReprtGrping.Visible = False
        lblGrp(0).Visible = False
        lblGrp(1).Visible = False
        lblGrp(2).Visible = False
        lblGrp(3).Visible = False
        comGroup(0).Visible = False
        comGroup(1).Visible = False
        comGroup(2).Visible = False
        comGroup(3).Visible = False
        chkAbsence.Visible = False
        chkInclTerm.Visible = False
    End If
        
    If (Me.Caption = "Costed Attendance Report" Or ReportSel = "COS") And glbCompSerial = "S/N - 2423W" Then   'Accellos Inc. - Ticket #26376
        'Preselect the Reason Codes
        clpCode(2).Text = "BT,VAC,UNSI,EXCU,POVT,REG"
    End If
    
    'Ticket #27095 - Opening up for ALL Attendance reports
    'Ticket #26272 Franks 01/08/2014 - begin
    'If ReportSel = "HIS" Then
    If ReportSel <> "EnviroServRpt" Then   'Ticket #22034
        chkInclTerm.Visible = True
    End If
    'End If
    'Ticket #26272 Franks 01/08/2014 - end
    
    optGrouping(0).Value = 1
End Sub

Private Sub Cri_Attend()
Dim EECri As String, OneSet%, X%

OneSet% = False
If Len(clpCode(2).Text) = 0 Then Exit Sub

EECri = EECri & "'" & Replace(clpCode(2), ",", "','") & "' "

HisSQL = HisSQL & " AND AD_REASON IN (" & EECri & ")"

EECri = "{HR_ATTENDANCE.AD_REASON} IN [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Upload()
Dim EECri As String

If cmbUpload.ListIndex > 0 Then
    If cmbUpload.ListIndex = 1 Then
        EECri = "{HR_ATTENDANCE.AD_UPLOAD}  = 'Y' "
        HisSQL = HisSQL & " AND AD_UPLOAD ='Y'"
    End If
    If cmbUpload.ListIndex = 2 Then
        EECri = "(isnull({HR_ATTENDANCE.AD_UPLOAD}) or {HR_ATTENDANCE.AD_UPLOAD}='N' )"
        HisSQL = HisSQL & " AND (AD_UPLOAD ='N' OR AD_UPLOAD IS NULL)"
    End If
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    glbstrSelCri = glbstrSelCri & EECri
End If
End Sub

Private Sub Cri_Discipline()
Dim EECri As String

If Len(clpCode(4)) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_DISCIPLINE} = '" & clpCode(4) & "' "
    HisSQL = HisSQL & " AND AD_DISCIPLINE='" & clpCode(4) & "' "
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

Private Sub Cri_ChargeCode()
Dim EECri As String

If Len(txtChargeCode.Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_CHRGCODE} = '" & txtChargeCode.Text & "' "
    HisSQL = HisSQL & " AND AD_CHRGCODE='" & txtChargeCode.Text & "' "
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

Private Sub Cri_Machine()
Dim EECri As String

If Len(clpCode(5).Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_MACHINE_NUM} = '" & clpCode(5).Text & "' "
    HisSQL = HisSQL & " AND AD_MACHINE_NUM='" & clpCode(5).Text & "' "
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

Private Sub Cri_Account()
Dim EECri As String

If Len(clpCode(11).Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_PROJECT_CODE} = '" & clpCode(11).Text & "' "
    HisSQL = HisSQL & " AND AD_PROJECT_CODE='" & clpCode(11).Text & "' "
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

Private Sub Cri_Absence()

If chkAbsence.Value Then
    If glbOracle Then
        glbstrSelCri = glbstrSelCri & " AND ({tblReason.TB_ABSENCE} <> 0) "
    Else
        glbstrSelCri = glbstrSelCri & " AND ({tblReason.TB_ABSENCE}) "
    End If
    HisSQL = HisSQL & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE<>0) "
End If
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri, strCd$
If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
        Case 0: strCd$ = "HREMP.ED_LOC"
        Case 1: strCd$ = IIf(glbMulti, IIf(ReportSel = "HIS" Or ReportSel = "COS", "HRATTWRK.AD_ORG", "HR_ATTENDANCE.AD_ORG"), "HREMP.ED_ORG")
        Case 6: strCd$ = "HREMP.ED_EMP"
        Case 7: strCd$ = "HRJOB.JB_GRPCD"
        Case 8: strCd$ = "HREMP.ED_SECTION"
        Case 9: strCd$ = IIf(glbCompSerial = "S/N - 2411W", "HR_ATTENDANCE.AD_REGION", "HREMP.ED_REGION")   'Ticket #24655
        Case 10: strCd$ = "HREMP.ED_ADMINBY"
    End Select
    CodeCri = "({" & strCd$ & "} in ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
    If glbMulti And intIdx% = 1 Then
        HisSQL = HisSQL & " AND AD_ORG IN ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "') "
    End If
    If intIdx% = 7 Then
        HisSQL = HisSQL & " AND AD_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(intIdx%).Text & "')) "
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
    'Hemu - Kumi Canada - Ticket #5507
    If glbCompSerial = "S/N - 2327W" And Left(clpDiv.Text, 1) = "-" Then
        DivCri = "({HREMP.ED_DIV} <> mid('" & clpDiv.Text & "',2))"
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

Private Sub Cri_DOW()
Dim EECri As String, OneSet%
Dim X%


For X% = 0 To 1
    If Len(medDOW(X%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then Exit Sub

        
'If Len(medDOW(0)) > 0 Then
'    If (Not glbSQL) And (Not glbOracle) Then
'        HisSQL = HisSQL & " AND (Weekday(AD_DOA) IN (" & medDOW(0)
'    Else
'        HisSQL = HisSQL & " AND (DatePart(dw,AD_DOA) IN (" & medDOW(0)
'    End If
'    If Len(medDOW(1)) > 0 Then
'        HisSQL = HisSQL & "," & medDOW(1) & "))"
'    Else
'        HisSQL = HisSQL & "))"
'    End If
'ElseIf Len(medDOW(1)) > 0 Then
'
'
'End If
        

For X% = 0 To 1
    If Len(medDOW(X%)) > 0 Then
        EECri = EECri & medDOW(X%)
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        EECri = EECri & " ,"
    End If
Next X%

If (Not glbSQL) And (Not glbOracle) Then
    HisSQL = HisSQL & " AND (Weekday(AD_DOA) IN (" & EECri & "))"
Else
    HisSQL = HisSQL & " AND (DatePart(dw,AD_DOA) IN (" & EECri & "))"
End If

EECri = "{@DOANumDay} in [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN [ " & getEmpnbr(elpEEID.Text) & "] "
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND (" & EECri & ") "
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

'Leeds and GrenVille -Mostafa
Private Sub Cri_CASE()
Dim CACri As String
Dim TableName As String

If ReportSel = "ATT" Then
    TableName = "HR_ATTENDANCE"
ElseIf ReportSel = "HIS" Or ReportSel = "COS" Then
    TableName = "HRATTWRK"
End If
If Len(TxtCases.Text) > 0 Then
    CACri = "{" & TableName & ".AD_PROJECT_CODE } IN ['" & Replace(TxtCases.Text, ",", "','") & "'] "
    
     If Len(HisSQL) > 0 Then
        HisSQL = HisSQL & " And " & TableName & ".AD_PROJECT_CODE in ('" & Replace(TxtCases.Text, ",", "','") & "') "
     Else
        HisSQL = HisSQL & " " & TableName & ".AD_PROJECT_CODE in ('" & Replace(TxtCases.Text, ",", "','") & "') "
     End If
End If

If Len(CACri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND (" & CACri & ") "
    Else
        glbstrSelCri = CACri
    End If
    glbiOneWhere = True
End If

End Sub

'Leeds and GrenVille -Mostafa
Private Sub Cri_Group()
Dim GRCri As String
Dim TableName As String

If ReportSel = "ATT" Then
    TableName = "HR_ATTENDANCE"
ElseIf ReportSel = "HIS" Or ReportSel = "COS" Then
    TableName = "HRATTWRK"
End If

If Len(cmbAttendGroup.Text) > 0 Then

    Dim SQLQ, gcode, codes
    gcode = Split(cmbAttendGroup.Text, " - ")
    SQLQ = "SELECT * FROM  HRATTGRP WHERE (AD_ATTENDANCE_GROUP = '" & gcode(0) & "') AND (AD_ATTENDANCE_TABLE = 'ADGP') AND (AD_ATTENDANCE_CODE_TABL = 'ADRE')"
    
    Dim rsCodes As New ADODB.Recordset
    rsCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    Dim ctr, I
    
    ctr = rsCodes.RecordCount
    I = 0
    Do Until rsCodes.EOF
        codes = codes & "'" & rsCodes("AD_ATTENDANCE_CODE") & "'"
        If I < (ctr - 1) Then
            codes = codes & ","
        End If
        I = I + 1
        rsCodes.MoveNext
    Loop
    rsCodes.Close
    Set rsCodes = Nothing
        
        If codes = "" Then
            codes = "''"
        End If
    GRCri = "{" & TableName & ".AD_REASON } IN [ " & codes & "] "
   ' MsgBox GRCri
     If Len(HisSQL) > 0 Then
        
        HisSQL = HisSQL & " And " & TableName & ".AD_REASON in (" & codes & ") "
     Else
        HisSQL = HisSQL & " " & TableName & ".AD_REASON in (" & codes & ") "
     End If
End If

If Len(GRCri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND (" & GRCri & ") "
    Else
        glbstrSelCri = GRCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_UseHireDate()
    Dim TempCri As String
    
    If chkUseHireDate.Visible = True Then
        If chkUseHireDate.Value = vbChecked Then
            TempCri = "({HR_ATTENDANCE.AD_DOA} >= {HREMP.ED_DOH})"
            HisSQL = HisSQL & " AND (AD_DOA >= HREMP.ED_DOH) "
            GoTo Cri_UseHireDate_Next
        End If
    End If
Exit Sub

Cri_UseHireDate_Next:
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
Dim X%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    HisSQL = HisSQL & " AND (AD_DOA Between "
    HisSQL = HisSQL & Date_SQL(dlpDateRange(0)) & "And "
    HisSQL = HisSQL & Date_SQL(dlpDateRange(1)) & ") "
    
    TempCri = "({HR_ATTENDANCE.AD_DOA} "
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
        TempCri = "({HR_ATTENDANCE.AD_DOA} "
        If X% = 0 Then
            TempCri = TempCri & " >= "
            HisSQL = HisSQL & "AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & ") "
        Else
            TempCri = TempCri & " <= "
            HisSQL = HisSQL & "AND (AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & " ) "
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

Private Sub Cri_Hours()
Dim lstEmpnbr, strWrkSQL As String
Dim rsAttWRK As New ADODB.Recordset

If Len(txtHours.Text) = 0 Then Exit Sub

If ReportSel = "HIS" Or ReportSel = "COS" Then
    strWrkSQL = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS HRS FROM HRATTWRK  WHERE AD_WRKEMP='" & glbUserID & "' GROUP BY AD_EMPNBR HAVING SUM(AD_HRS) >= " & Val(txtHours.Text)
    rsAttWRK.Open strWrkSQL, gdbAdoIhr001W, adOpenForwardOnly
Else
    strWrkSQL = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS HRS FROM HR_ATTENDANCE WHERE " & HisSQL & " GROUP BY AD_EMPNBR HAVING SUM(AD_HRS) >= " & Val(txtHours.Text)
    rsAttWRK.Open strWrkSQL, gdbAdoIhr001, adOpenForwardOnly
End If
lstEmpnbr = "-1"
If Not rsAttWRK.EOF Then
    rsAttWRK.MoveFirst
    
    Do While Not rsAttWRK.EOF
        'lstEmpnbr Employee #s to form the selection formula for the report
        lstEmpnbr = lstEmpnbr & "," & rsAttWRK("EMPNBR")
        rsAttWRK.MoveNext
    Loop
End If
rsAttWRK.Close
            
    
If lstEmpnbr <> "" Then
    glbstrSelCri = glbstrSelCri & " AND ({HREMP.ED_EMPNBR} IN [ " & lstEmpnbr & "])"
    HisSQL = HisSQL & " AND (ED_EMPNBR IN (" & lstEmpnbr & "))"
End If

 
End Sub

Private Sub Cri_HoursLess()
Dim lstEmpnbr, strWrkSQL As String
Dim rsAttWRK As New ADODB.Recordset

If Len(txtHoursLess.Text) = 0 Then Exit Sub

If ReportSel = "HIS" Or ReportSel = "COS" Then
    strWrkSQL = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS HRS FROM HRATTWRK  WHERE AD_WRKEMP='" & glbUserID & "' GROUP BY AD_EMPNBR HAVING SUM(AD_HRS) <= " & Val(txtHoursLess.Text)
    rsAttWRK.Open strWrkSQL, gdbAdoIhr001W, adOpenForwardOnly
Else
    strWrkSQL = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS HRS FROM HR_ATTENDANCE WHERE " & HisSQL & " GROUP BY AD_EMPNBR HAVING SUM(AD_HRS) <= " & Val(txtHoursLess.Text)
    rsAttWRK.Open strWrkSQL, gdbAdoIhr001, adOpenForwardOnly
End If
lstEmpnbr = "-1"
If Not rsAttWRK.EOF Then
    rsAttWRK.MoveFirst
    
    Do While Not rsAttWRK.EOF
        'lstEmpnbr Employee #s to form the selection formula for the report
        lstEmpnbr = lstEmpnbr & "," & rsAttWRK("EMPNBR")
        rsAttWRK.MoveNext
    Loop
End If
rsAttWRK.Close
            
    
If lstEmpnbr <> "" Then
    glbstrSelCri = glbstrSelCri & " AND ({HREMP.ED_EMPNBR} IN [ " & lstEmpnbr & "])"
    HisSQL = HisSQL & " AND (ED_EMPNBR IN (" & lstEmpnbr & "))"
End If

End Sub

Private Sub Cri_Incident()
Dim EECri As String

If Len(txtIncident.Text) > 0 Then
    HisSQL = HisSQL & " AND AD_INCID " & IIf(txtIncident = "Y", "<>", "=") & "0 "
    
    If glbOracle Then
        EECri = "{HR_ATTENDANCE.AD_INCID}" & IIf(txtIncident = "Y", "<>", "=") & "0"
    Else
        If txtIncident = "Y" Then
            EECri = "{HR_ATTENDANCE.AD_INCID}"
        Else
            EECri = " NOT {HR_ATTENDANCE.AD_INCID}"
        End If
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

Private Sub Cri_Incentive()
Dim EECri As String

If Len(txtIncentive.Text) > 0 Then
    HisSQL = HisSQL & " AND AD_INDICATOR " & IIf(txtIncentive = "Y", "<>", "=") & "0 "
    
    If glbOracle Then
        EECri = "{HR_ATTENDANCE.AD_INDICATOR}" & IIf(txtIncentive = "Y", "<>", "=") & "0"
    Else
        If txtIncentive = "Y" Then
            EECri = "{HR_ATTENDANCE.AD_INDICATOR}"
        Else
            EECri = " NOT {HR_ATTENDANCE.AD_INDICATOR}"
        End If
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

Private Sub Cri_Seniority()
Dim EECri As String

If Len(txtSeniority.Text) > 0 Then
    HisSQL = HisSQL & " AND AD_SEN " & IIf(txtSeniority = "Y", "<>", "=") & "0 "
    
    If glbOracle Then
        EECri = "{HR_ATTENDANCE.AD_SEN}" & IIf(txtSeniority = "Y", "<>", "=") & "0"
    Else
        If txtSeniority = "Y" Then
            EECri = "{HR_ATTENDANCE.AD_SEN}"
        Else
            EECri = " NOT {HR_ATTENDANCE.AD_SEN}"
        End If
    End If
End If

'If Len(txtSeniority.Text) > 0 Then
'    If glbOracle Then
'        If ReportSel = "HIS" Then
'            If txtSeniority = "Y" Then
'                EECri = "{HRATTWRK.AD_SEN}<>0"
'            Else
'                EECri = " {HRATTWRK.AD_SEN}=0"
'            End If
'        Else
'            If txtSeniority = "Y" Then
'                EECri = "{HR_ATTENDANCE.AD_SEN}<>0"
'            Else
'                EECri = " {HR_ATTENDANCE.AD_SEN}=0"
'            End If
'        End If
'    Else
'        If ReportSel = "HIS" Then
'            If txtSeniority = "Y" Then
'                EECri = "{HRATTWRK.AD_SEN}"
'            Else
'                EECri = " NOT {HRATTWRK.AD_SEN}"
'            End If
'        Else
'            If txtSeniority = "Y" Then
'                EECri = "{HR_ATTENDANCE.AD_SEN}"
'            Else
'                EECri = " NOT {HR_ATTENDANCE.AD_SEN}"
'            End If
'        End If
'    End If
'End If


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
Dim X%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

flgUseLabelMst = True
glbiOneWhere = False
glbstrSelCri = ""


'--------hremp relates start
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Code(0)
Call Cri_Code(6)
Call Cri_Code(8)
Call Cri_Code(9)
Call Cri_Code(10)

Call Cri_PT
Call Cri_EE
If glbMulti Then
    HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    '--------hremp relates   end
    '--------attendance relates start
    Call Cri_Code(1)
Else
    Call Cri_Code(1)
    
    If ReportSel = "COS" And glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Limit the user has access to these employees salary information
        If Len(glbNoAccessGrp) <> 0 Then
            HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
            If Len(HisSQL) > 0 Then
                HisSQL = HisSQL & " AND (" & glbNoAccessGrp & ") "
            Else
                HisSQL = HisSQL & " (" & glbNoAccessGrp & ") "
            End If
            'HisSQL = Replace(Replace(HisSQL, "HREMP.ED_", "ED_"), "ED_", "HREMP.ED_")
            'HisSQL = Replace(HisSQL, "(HREMP.", "({HREMP.")
            'HisSQL = Replace(HisSQL, " HREMP.", " {HREMP.")
            'HisSQL = Replace(HisSQL, " =", "} =")
            HisSQL = HisSQL & ")"
        Else
            HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
        End If
    Else
        HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
        '--------hremp relates   end
        '--------attendance relates start
    End If
End If

Call Cri_ChargeCode
Call Cri_Machine
Call Cri_Account
If glbWFC And glbPlantCode = "WHBY" Then
    Call Cri_Discipline
End If
Call Cri_Attend
Call Cri_FTDates
Call Cri_Upload
Call Cri_DOW
Call Cri_Shift
Call Cri_Sup
Call Cri_Incident
Call Cri_Incentive  'Ticket #15276
Call Cri_Seniority
'--------attendance relates end
'--------other
Call Cri_Job ' this only for multi position
Call Cri_Code(7)
Call Cri_Absence
'--------other end
'Ticket #21631 - Get Attendance records from Hire Date
Call Cri_UseHireDate

If ReportSel = "HIS" Or ReportSel = "COS" Then
    If optGrouping(2) And glbCompSerial = "S/N - 2411W" Then   'Ticket #24655 - WDGPHU
        Call WDGPHU_PayrollAttendance
    Else
        'transffer data to working table
        'Ticket #20039 & Ticket #20047
        Call SELATTWRK1
    End If
    'The working table contains the selection criteria already, do not need pass the selection to reports.
    glbstrSelCri = " {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
Else
    If chkInclTerm.Value Then 'Ticket #27095
        'Ticket #27095 - Mainly for Include Terminated employees
        Call SELATTWRK1
        
        'The working table contains the selection criteria already, do not need pass the selection to reports.
        glbstrSelCri = " {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    End If
End If

' Cri_Hours must to be called after SELATTWRK
Call Cri_Hours
'Ticket #13035 Frank Jun 25, 2007
Call Cri_HoursLess

'commented by Bryan Ticket#12094 Updateing the salary is being done during SelAttWrk.
' Update_Att_Salary one must to be called after SELATTWRK
'If ReportSel = "COS" And chkCostHis.Value = False Then
'    If glbCompSerial = "S/N - 2192W" Then  ' county essex
'        'do not update the salary field using sh_salary
'    Else
'        Call Update_Att_Salary
'    End If
'End If

' report name


'Leeds and Grenville - Mostafa
If glbCompSerial = "S/N - 2233W" Then
    Call Cri_Group
    Call Cri_CASE
End If

If ReportSel = "ATT" Then
    If optGrouping(0) Then
        If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2236_rzattend.rpt"
            flgUseLabelMst = False
        ElseIf chkInclTerm.Value Then 'Ticket #27095
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattendT.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattend.rpt"
        End If
    ElseIf optGrouping(3) Then  'Release 8.1
        If chkInclTerm.Value Then 'Ticket #27095
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattensTTot.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattensTot.rpt"
        End If
    Else
        If chkInclTerm.Value Then 'Ticket #27095
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattensT.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattens.rpt"
        End If
    End If
    Me.vbxCrystal.WindowTitle = "Attendance Report"
End If
If ReportSel = "HIS" Then
    If optGrouping(0) Then
        If glbCompSerial = "S/N - 2233W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2233_rzattdhd.rpt"
            flgUseLabelMst = False
        ElseIf chkInclTerm.Value Then 'Ticket #26272 Franks 03/31/2015
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhdT.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhd.rpt"
        End If
    ElseIf optGrouping(3) Then  'Release 8.1
        If chkInclTerm.Value Then 'Ticket #26272 Franks 03/31/2015
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhsTTot.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhsTot.rpt"
        End If
    Else
        If glbCompSerial = "S/N - 2233W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2233_rzattdhs.rpt"
            flgUseLabelMst = False
        ElseIf chkInclTerm.Value Then 'Ticket #26272 Franks 03/31/2015
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhsT.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhs.rpt"
        End If
    End If
    Me.vbxCrystal.WindowTitle = "Attendance Report Including Historical Data"
End If

If ReportSel = "COS" Then
    If optGrouping(0) Then
        If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then   'laura 03/09/98
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsattdct.rpt"
            flgUseLabelMst = False
        ElseIf chkCostHis Then
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattchtT.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattcht.rpt"
            End If
        Else
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdctT.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdct.rpt"
            End If
        End If
    ElseIf optGrouping(3) Then  'Release 8.1
        If chkCostHis Then
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattchsTTot.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattchsTot.rpt"
            End If
        Else
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdcsTTot.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdcsTot.rpt"
            End If
        End If
    Else
        If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then   'laura 03/09/98
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsattdcs.rpt"
            flgUseLabelMst = False
        ElseIf optGrouping(2) And glbCompSerial = "S/N - 2411W" Then   'Ticket #24655 - WDGPHU
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2411_PAttend.rpt"
            Me.vbxCrystal.Formulas(6) = "title='Costed Attendance Report for Payroll for the Pay Period #" & txtWeek.Text & "'"
            Me.vbxCrystal.Formulas(7) = "lblAttRegion='" & lStr("Region") & "'"
            Me.vbxCrystal.Formulas(10) = "ShowComments=" & (chkComments <> 0)
        ElseIf chkCostHis Then
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattchsT.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattchs.rpt"
            End If
        Else
            If chkInclTerm.Value Then 'Ticket #27095
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdcsT.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdcs.rpt"
            End If
        End If
    End If
    Me.vbxCrystal.WindowTitle = "Costed Attendance Report"
End If

If ReportSel = "PerDayRpt" Then
    If optGrouping(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattpday.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattpdas.rpt"
    End If
    Me.vbxCrystal.WindowTitle = "Attendance Report"
End If

'Ticket #22034 - County of Wellington
If ReportSel = "EnviroServRpt" Then
    Call Wellington_Terrace_Attendance_Report
    GoTo Skip
End If

' set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri   'Hemu - By default, has to be there even when you remove Hemu's changes
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    
    If ReportSel = "HIS" Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 6
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
    ElseIf ReportSel = "COS" Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        If chkCostHis Then
            For X% = 0 To 9
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
            Me.vbxCrystal.DataFiles(3) = glbIHRDBW
        
        Else
            For X% = 0 To 11
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            Next X%
            Me.vbxCrystal.DataFiles(4) = glbIHRDBW
        
        End If
    Else
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
End If

Skip:
Cri_SetAll = True

DoEvents

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub Cri_Shift()
Dim EECri As String

If Len(txtShift.Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_SHIFT} = '" & txtShift.Text & "' " 'laura dec 16, 1997
    HisSQL = HisSQL & " AND AD_SHIFT ='" & txtShift.Text & "'"
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$, DOA
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3
Dim strSMonth$, I

z% = 0
Cri_Sorts = 0


'Set the report labels as per the label master - if not custom report
If flgUseLabelMst = True Then
    Call setRptLabel(Me, 4)
End If

If ReportSel <> "HIS" Then 'laura nov 25, 1997
  If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
      strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
      Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
      strSMonth$ = DateDiff("m", dlpDateRange(0).Text, dlpDateRange(1).Text) + 1
      Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$

  Else
      strSFormat$ = "No date entered"
      Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
      strSMonth$ = "1"
      Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$

  End If
Else
  If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
      strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
      Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

  Else
      strSFormat$ = "No date entered"
      Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

  End If
End If

'Frank 04/02/04 Ticket# 5937
If glbWFC And glbPlantCode = "WHBY" Then
    If ReportSel = "COS" Then
        If optGrouping(0) And chkCostHis Then
            Me.vbxCrystal.Formulas(4) = "whitby = true"
        End If
    Else
        Me.vbxCrystal.Formulas(4) = "whitby = true"
    End If
End If

'Ticket #24066 - Town of St. Marys - Show GL instead of Charge Code
If glbCompSerial = "S/N - 2384W" Then
    If ReportSel = "ATT" And optGrouping(0) Then
        Me.vbxCrystal.Formulas(4) = "stmarys = true"
    End If
End If

'Hemu - 06/02/2003 Begin - Hours in the Report Title
If Len(txtHours.Text) > 0 Then
    Me.vbxCrystal.Formulas(11) = "Hours = '" & " Hours >= " & txtHours.Text & "'"
End If
'Hemu - 06/02/2003 End - Hours in the Report Title

If comGroup(0).Text = "Position" Then
    grpField$ = getEGroup("Position Code")
Else
    'Ticket #29984 - Option to use Employee Master data instead of Attendance
    If comGroup(0).Text = lStr("G/L #") And chkUseEmpMst.Value Then
        grpField$ = "{HRGL_1.GL_DESCR}"
    Else
        grpField$ = getEGroup(comGroup(0).Text)
    End If
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21854 Franks 04/09/2012, they need code instead of desc
    If comGroup(0).Text = lStr("Administered By") Then
        grpField$ = "{HREMP.ED_ADMINBY}"
    End If
End If
If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"

If ReportSel = "HIS" Or ReportSel = "COS" Then
    DOA = "{HRATTWRK.AD_DOA}"
    If comGroup(0).Text = "Shift" Then grpField$ = "{HRATTWRK.AD_SHIFT}"
Else
    If chkInclTerm.Value Then 'Ticket #27095
        DOA = "{HRATTWRK.AD_DOA}"
        If comGroup(0).Text = "Shift" Then grpField$ = "{HRATTWRK.AD_SHIFT}"
    Else
        DOA = "{HR_ATTENDANCE.AD_DOA}"
        If comGroup(0).Text = "Shift" Then grpField$ = "{HR_ATTENDANCE.AD_SHIFT}"
    End If
End If

'Added by Bryan 30/May/05 Ticket#10902
If comGroup(0).Text = lStr("AttSupervisor") Then
    If ReportSel = "ATT" Then
        If chkSup.Value = vbUnchecked Then
            grpField$ = "{@fldADSuper}"
        Else
            grpField$ = "{@fldJBSuper}"
        End If
    Else
        If chkSup.Value = vbUnchecked Then
            grpField$ = "{@fldADSuper}"
        Else
            grpField$ = "{@fldJBSuper}"
        End If
    End If
End If


If frmRAttend.Caption = "Costed Attendance Report" Then 'Frank 4/24/2000
    SavGrp1 = grpField$
    If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'If GrpIdx% < 6 Then
      Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
      Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
    Else
      Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
      Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
      Me.vbxCrystal.SectionFormat(z%) = "GF1;F;X;X;X;X;X;X"
      z% = z% + 1
    End If
Else
    SavGrp1 = grpField$
    If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'If GrpIdx% < 5 Then
        If glbCompSerial = "S/N - 2327W" And comGroup(0).Text = "Employee Name" Then
            Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Associate Name'"
        Else
            Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
        End If
      Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
    Else
      Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
      Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
      Me.vbxCrystal.SectionFormat(z%) = "GF1;F;X;X;X;X;X;X"
      z% = z% + 1
    End If
End If

Me.vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"


' SECOND primary grouping ------------------------- **
GrpIdx% = comGroup(1).ListIndex
Me.vbxCrystal.Formulas(5) = "DESCGROUP2H = '" & comGroup(1).Text & "'"      'Release 8.1
If frmRAttend.Caption = "Attendance Report" Then 'Ticket #21654 Franks 03/22/2012
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblReason.TB_DESC}"
        Case 2: grpField$ = "{HR_PROJECT_CODE.DESCRIPTION}"     'Ticket #26118 - Add Account Code to G1 and G2
        'Case 2: grpField$ = "{tblSec.TB_DESC}"
        'Samuel Ticket #21854 Franks 04/09/2012, they need code instead of desc
        Case 3: grpField$ = IIf(glbCompSerial = "S/N - 2382W", "{HREMP.ED_SECTION}", "{tblSec.TB_DESC}")
        Case 4: grpField$ = getEGroup("Position Code")  '(comGroup(1).Text)      'Release 8.1
        Case 5: grpField$ = SavGrp1
    End Select
Else
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblReason.TB_DESC}"
        Case 2: grpField$ = "{HR_PROJECT_CODE.DESCRIPTION}"     'Ticket #26118 - Add Account Code to G1 and G2
        Case 3: grpField$ = getEGroup("Position Code")  'getEGroup(comGroup(1).Text)  'Ticket #29852 Franks 02/24/2017
        Case 4: grpField$ = SavGrp1
    End Select
End If
SavGrp2 = grpField$

'Ticket #27095
If chkInclTerm.Value Then
    Me.vbxCrystal.GroupCondition(2) = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
Else
    Me.vbxCrystal.GroupCondition(1) = "GROUP2;" & grpField$ & ";ANYCHANGE;A"
End If


If comGroup(1).ListIndex > 1 Then
    If glbCompSerial = "S/N - 2382W" And (comGroup(1).Text = lStr("Section")) Then   'Samuel Ticket #21854 Franks 04/09/2012, disp group 2 total
        Me.vbxCrystal.SectionFormat(z%) = "GF2;T;X;X;X;X;X;X"
    Else
        'Ticket #27095
        If chkInclTerm.Value Then
            If (frmRAttend.Caption = "Attendance Report" And comGroup(1).ListIndex = 5) Or (frmRAttend.Caption <> "Attendance Report" And comGroup(1).ListIndex = 4) Then
                Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
            Else
                'Release 8.1
                Me.vbxCrystal.SectionFormat(z%) = "GH3;T;X;X;X;X;X;X"
                z% = z% + 1
                Me.vbxCrystal.SectionFormat(z%) = "GF3;T;X;X;X;X;X;X"
            End If
        Else
            If (frmRAttend.Caption = "Attendance Report" And comGroup(1).ListIndex = 5) Or (frmRAttend.Caption <> "Attendance Report" And comGroup(1).ListIndex = 4) Then
                Me.vbxCrystal.SectionFormat(z%) = "GF2;F;X;X;X;X;X;X"
            Else
                'Release 8.1
                Me.vbxCrystal.SectionFormat(z%) = "GH2;T;X;X;X;X;X;X"
                z% = z% + 1
                Me.vbxCrystal.SectionFormat(z%) = "GF2;T;X;X;X;X;X;X"
            End If
        End If
    End If
Else
    'Ticket #27095
    If chkInclTerm.Value Then
        Me.vbxCrystal.SectionFormat(z%) = "GF3;T;X;X;X;X;X;X"
    Else
        Me.vbxCrystal.SectionFormat(z%) = "GF2;T;X;X;X;X;X;X"
    End If
End If

z% = z% + 1

If Not (ReportSel = "PerDayRpt") And Not (ReportSel = "COS" And optGrouping(2) And glbCompSerial = "S/N - 2411W") Then
' 3TH primary grouping ------------------------- **8
    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblReason.TB_DESC}"
        Case 1: grpField$ = SavGrp2
    End Select
    If optGrouping(0) Then
        'grouping 3 is for ED_EMPNBR
    
    '    Me.vbxCrystal.GroupCondition(2) = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
    '    If comGroup(2).ListIndex > 0 Then
    '        Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
    '    Else
    '        Me.vbxCrystal.SectionFormat(z%) = "GF3;T;X;X;X;X;X;X"
    '    End If
    
        'Ticket #27095
        If chkInclTerm.Value Then
            Me.vbxCrystal.GroupCondition(4) = "GROUP5;" & grpField$ & ";ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(3) = "GROUP4;" & grpField$ & ";ANYCHANGE;A"
        End If
        
        If comGroup(2).ListIndex > 0 Then
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.SectionFormat(z% + 1) = "GF5;F;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.SectionFormat(z% + 1) = "GF4;F;X;X;X;X;X;X"
            End If
        Else
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.SectionFormat(z% + 1) = "GF5;T;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.SectionFormat(z% + 1) = "GF4;T;X;X;X;X;X;X"
            End If
        End If
        Me.vbxCrystal.Formulas(10) = "ShowComments=" & (chkComments <> 0)
    Else
        If comGroup(2).ListIndex > 0 Then
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.SectionFormat(z%) = "GF4;F;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
            End If
        Else
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.SectionFormat(z%) = "GF4;T;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.SectionFormat(z%) = "GF3;T;X;X;X;X;X;X"
            End If
        End If
    End If
    z% = z% + 1
    ' END ---------------------------------------------- ***
    
    ' SUMMARY ------------------------------------------ ***
    If optGrouping(1) Then
        'Ticket #27095
        If chkInclTerm.Value Then
            Me.vbxCrystal.SectionFormat(z%) = "GF4;F;X;X;X;X;X;X"
        Else
            Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
        End If
        z% = z% + 1
        'MsgBox Str(comGroup(1).ListIndex)
        If comGroup(1).ListIndex = 0 Then
            
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;T;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(4) = "GROUP5;{HREMP.ED_COMPNO};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(4) = "GF5;F;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.GroupCondition(2) = "GROUP3;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(2) = "GF3;T;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{HREMP.ED_COMPNO};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
            End If
        End If
        If comGroup(1).ListIndex = 1 Then
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{tblReason.TB_DESC};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;T;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(4) = "GROUP5;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(4) = "GF5;F;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.GroupCondition(2) = "GROUP3;{tblReason.TB_DESC};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(2) = "GF3;T;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
            End If
        End If
        If comGroup(1).ListIndex = 2 Then
            If glbCompSerial = "S/N - 2382W" And (comGroup(1).Text = lStr("Section")) Then 'Samuel Ticket #21854 Franks 04/09/2012,
                'disp group 2 total
            Else
                'Not sure why for Summary Report with Account Code Grouping 2 - the Group Footer for Account Code Total is being hidden
                ''Ticket #27095
                'If chkInclTerm.Value Then
                '    Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HREMP.ED_COMPNO};ANYCHANGE;A"
                '    Me.vbxCrystal.SectionFormat(2) = "GF3;F;X;X;X;X;X;X"
                'Else
                '    Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HREMP.ED_COMPNO};ANYCHANGE;A"
                '    Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
                'End If
            End If
            
            'Ticket #27095
            If chkInclTerm.Value Then
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{HREMP.ED_COMPNO};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(4) = "GROUP5;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(4) = "GF5;F;X;X;X;X;X;X"
            Else
                Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HREMP.ED_COMPNO};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(2) = "GF3;F;X;X;X;X;X;X"
        
                Me.vbxCrystal.GroupCondition(3) = "GROUP4;{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.SectionFormat(3) = "GF4;F;X;X;X;X;X;X"
            End If
        End If
    
    End If
End If

If (ReportSel = "COS" And optGrouping(2) And glbCompSerial = "S/N - 2411W") Then
    If comGroup(1).Text = "(none)" And comGroup(2).Text = "(none)" Then Me.vbxCrystal.SectionFormat(z%) = "GF3;F;X;X;X;X;X;X"
End If

If (ReportSel = "PerDayRpt") Then
    For I = 1 To 6
        Me.vbxCrystal.Formulas(I + 20) = (fDis(I))
    Next I
End If
    

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpSUP(1).Text) > 0 Then
    If ReportSel = "ATT" Then
        If chkSup.Value = vbUnchecked Then
            EECri = "{HR_ATTENDANCE.AD_SUPER} IN [" & getEmpnbr(elpSUP(1).Text) & "] "
        Else 'Added by Bryan 28/Dec/05 Ticket#10058
            EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpSUP(1).Text) & "] "
            EECri = EECri & " AND {HR_JOB_HISTORY.JH_CURRENT} = True"   'Ticket #21731
        End If
        'If glbSQL Or glbOracle Then EECri = Replace(Replace(EECri, "[", "("), "]", ")")
    Else
        If chkSup.Value = vbUnchecked Then
            HisSQL = HisSQL & "AND (AD_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ")) "
        Else 'Added by Bryan 28/Dec/05 Ticket#10058
            HisSQL = HisSQL & "AND (JH_REPTAU IN (" & getEmpnbr(elpSUP(1).Text) & ")) "
            HisSQL = HisSQL & "AND JH_CURRENT <> 0 "   'Ticket #21731
        End If
        'If glbSQL Or glbOracle Then HisSQL = Replace(Replace(HisSQL, "[", "("), "]", ")")
    End If
Else
    'Ticket #21731
    If ReportSel = "ATT" Then
        If chkSup.Value = vbChecked Then
            EECri = "{HR_JOB_HISTORY.JH_CURRENT} = True"
        End If
        'If glbSQL Or glbOracle Then EECri = Replace(Replace(EECri, "[", "("), "]", ")")
    Else
        If chkSup.Value = vbChecked Then
            HisSQL = HisSQL & "AND JH_CURRENT <> 0 "   'Ticket #21731
        End If
        'If glbSQL Or glbOracle Then HisSQL = Replace(Replace(HisSQL, "[", "("), "]", ")")
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

Private Function CriCheck()
Dim X%

CriCheck = False
If glbLinamar And ReportSel = "ATT" Then
    If optGrouping(0) Then
        If Len(Trim(clpDiv.Text)) = 0 Then
            MsgBox lStr("Division is a required field")
             clpDiv.SetFocus
            Exit Function
        End If
    End If
    For X% = 0 To 1
        If Len(dlpDateRange(X%).Text) = 0 Then
            MsgBox IIf(X% = 0, "From", "To") & " is a required field"
            dlpDateRange(X%).Text = ""
            dlpDateRange(X%).SetFocus
            Exit Function
        End If
    Next X%
End If

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'Hemu - Kumi Canada - Ticket #5507
    If glbCompSerial = "S/N - 2327W" And Left(clpDiv.Text, 1) = "-" Then
        'do nothing
    Else
        'MsgBox lStr("If Division Entered - it must be known")
        'clpDiv.SetFocus
        Exit Function
    End If
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If


If elpEEID.Caption = "Enter Valid Employee #" Then
    MsgBox "If Employee Entered - they must exist"
    elpSUP(1).SetFocus
    Exit Function
End If

For X% = 0 To 10
    If X% < 1 Or X% > 5 Then
        If Not clpCode(X).ListChecker Then Exit Function
    End If
Next X%

If Not clpCode(2).CheckList Then
    clpCode(2).SetFocus
    Exit Function
End If

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

For X% = 0 To 1
 If Len(medDOW(X%)) > 0 Then
    If medDOW(X%) < 0 Or medDOW(X%) > 7 Then
        MsgBox "Day of Week must be between 1 and 7"
        medDOW(X%).SetFocus
        Exit Function
    End If
  End If
Next X%
 

If Len(txtIncident) > 0 Then
    If txtIncident <> "Y" And txtIncident <> "N" Then
        MsgBox "Incident must be Y/N or blank"
        txtIncident.SetFocus
        Exit Function
    End If
End If

If Len(txtIncentive) > 0 Then
    If txtIncentive <> "Y" And txtIncentive <> "N" Then
        If glbCompSerial = "S/N - 2214W" Then 'Casey House  - Ticket #15276
            MsgBox "HOOPP must be Y/N or blank"
        Else
            MsgBox "Incentive must be Y/N or blank"
        End If
        txtIncentive.SetFocus
        Exit Function
    End If
End If

If Len(txtSeniority) > 0 Then
    If UCase(txtSeniority) <> "Y" And UCase(txtSeniority) <> "N" Then
        MsgBox "Seniority must be Y/N or blank"
        txtSeniority.SetFocus
        Exit Function
    End If
End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

If Len(txtHours) > 0 Then
    If Not IsNumeric(txtHours) Then
        MsgBox "Hours must be number"
        txtHours.SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then Exit Function
If Not clpCode(1).ListChecker Then Exit Function

If glbCompSerial = "S/N - 2214W" Then
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox "If Attendance-Department is entered it must be valid"
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(3)) > 0 Then
        If clpCode(3).Caption = "Unassigned" Then
            MsgBox "Invalid Attendance-Fund code"
            clpCode(3).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2396W" Then  'Ticket #17323 - Oshawa CHC
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox "If G/L # is entered it must be valid"
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
End If

If glbWFC And glbPlantCode = "WHBY" Then
    If Len(clpCode(4)) > 0 Then
        If clpCode(4).Caption = "Unassigned" Then
            MsgBox "If code entered it must be known"
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
End If

If ReportSel = "PerDayRpt" Then 'Ticket #15133
    Call GetDisData
End If

If ReportSel = "COS" Then
    'Ticket #24655 - WDGPHU
    If glbCompSerial = "S/N - 2411W" And optGrouping(2) Then
        If Len(txtYear) = 0 Or Len(txtWeek) = 0 Then
            MsgBox "Year and the Pay Period # is mandatory"
            txtYear.SetFocus
            Exit Function
        Else
            If Val(txtYear) > Year(Date) + 100 Or Val(txtYear) < Year(Date) - 100 Then
                MsgBox "Year is not a valid year"
                txtYear.SetFocus
                Exit Function
            End If
        End If
    End If
End If

CriCheck = True
End Function

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()

TxtCases.Text = glbCaseFiles

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMRATTEND"

Call comGrpLoad

Call setRptCaption(Me)
Call setCaption(lblAccount)
Call setCaption(lblMachine)
Call setCaption(lblChargeCode)
Call setCaption(lblEENum(1))
Call setCaption(lblHours)
Call setCaption(lblHoursLess)
Call setCaption(lblShift)
chkUseHireDate.Caption = lStr("Use Original Hire Date")

If lblEENum(1).Caption = "AttSupervisor" Then lblEENum(1).Caption = "Supervisor"
lblAttCodes.Caption = lStr("Reason")
lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

If glbCompSerial = "S/N - 2227W" Then clpCode(9).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbLinamar Then
    clpCode(9).MaxLength = 8
    lblDiv.FontBold = True
    lblFromTo.FontBold = True
End If

If glbMulti Then
    lblUnion.ForeColor = &HC000C0
    lblJOB.Visible = True
    clpJOB.Visible = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call INI_Controls(Me)

cmbUpload.AddItem "All"
cmbUpload.AddItem "Yes"
cmbUpload.AddItem "No"
cmbUpload.ListIndex = 0

'Casey House Ticket #5522
If glbCompSerial = "S/N - 2214W" Then
    lblChargeCode.Caption = "Attendance-Dept"
    lblShift.Caption = "Attendance-Fund"
    txtChargeCode.Visible = False
    clpChrgCode.Visible = True
    txtShift.MaxLength = 4
    txtShift.Visible = False
    clpCode(3).Visible = True
End If

If glbCompSerial = "S/N - 2241W" Then
    clpChrgCode.Visible = True
    clpChrgCode.MaxLength = 7
    clpChrgCode.LookupType = Department
    txtChargeCode.Visible = False
    lblChargeCode.Caption = "Department"

End If

If glbWFC And glbPlantCode = "WHBY" Then
    lblChargeCode.Visible = False
    txtChargeCode.Visible = False
    lblTitle(9).Visible = True
    clpCode(4).Visible = True
    lblTitle(9).Top = lblChargeCode.Top '2760
    clpCode(4).Top = txtChargeCode.Top ' 2760
End If

If glbCompSerial = "S/N - 2192W" Then
    'charge code link to gl account
    lblChargeCode = lStr("G/L #")
    txtChargeCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
End If

If glbCompSerial = "S/N - 2376W" Then  'Ticket #16589 AFN
    'charge code link to gl account
    lblChargeCode = lStr("G/L #")
    txtChargeCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
End If

If glbCompSerial = "S/N - 2396W" Then  'Ticket #17323 - Oshawa CHC
    'Charge Code link to GL Account
    lblChargeCode = lStr("G/L #")
    txtChargeCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
    clpChrgCode.Tag = lStr("00-Enter G/L #")
End If


If glbCompSerial = "S/N - 2233W" Then
    lblAccount.Visible = False
    Me.clpCode(11).Visible = False
End If

If glbCompSerial = "S/N - 2279W" Then   'Friesens Corporation - Ticket #16189
    'Hide Machine Code fields
    clpCode(5).Visible = False
    lblMachine.Visible = False
End If


If glbVadim Then
    lblMachine = "Equipment"
    clpCode(5).TABLTitle = "Equipment"
    clpCode(5).Tag = "Equipment #"
End If

If glbCompSerial = "S/N - 2214W" Then 'Casey House  - Ticket #15276
    lblIncentive.Caption = "HOOPP"
    txtIncentive.Tag = "HOOPP - Y or N"
End If

If glbBurlTech Then
    'lblSen1.Caption = "Unexcused"
    lblSen1.Caption = "Excused"
End If
If glbLambton Then
    txtShift.MaxLength = 4
End If

If glbCompSerial = "S/N - 2233W" Then 'leeds and grenville only - Mostafa
    Call LoadGroupCodes
    lblCodes.Visible = True
    LblCases.Visible = True
    cmbAttendGroup.Visible = True
    TxtCases.Visible = True
Else
    lblCodes.Visible = False
    LblCases.Visible = False
    cmbAttendGroup.Visible = False
    TxtCases.Visible = False
End If

'Ticket #21840 - County of Wellington
If glbCompSerial = "S/N - 2262W" Then
    chkUseHireDate.Value = vbChecked
End If

'Ticket #24655 - WDGPHU
If ReportSel = "COS" And glbCompSerial = "S/N - 2411W" Then
    optGrouping(2).Visible = True
    chkComments.Left = 1845
    chkComments.Top = 6600
Else
    optGrouping(2).Visible = False
    chkComments.Left = 6240
    chkComments.Top = 6340
End If

If (Me.Caption = "Costed Attendance Report" Or ReportSel = "COS") And glbCompSerial = "S/N - 2423W" Then 'Accellos Inc. - Ticket #26376
    'Preselect the Reason Codes
    clpCode(2).Text = "BT,VAC,UNSI,EXCU,POVT,REG"
End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub LoadGroupCodes()
    Dim rsGCodes As New ADODB.Recordset
    Dim SQLQ
    Dim I
    I = 0
    
    SQLQ = "select TB_KEY, TB_DESC from hrtabl where tb_name ='ADGP'"
 
    rsGCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    
     'cmbAttendGroup.DataField = "TB_KEY"
     'cmbAttendGroup.DataMember = "TB_DESC"
     
     Do Until rsGCodes.EOF
         
         cmbAttendGroup.AddItem (rsGCodes("TB_KEY"))
         
        cmbAttendGroup.List(I) = rsGCodes("TB_KEY") & " - " & rsGCodes("TB_DESC")
         rsGCodes.MoveNext
         I = I + 1
         'MDIMain.panHelp(0).Caption = I & " Records Processed"
     Loop
    ' gdbAdoIhr001W.CommitTrans
     rsGCodes.Close
     Set rsGCodes = Nothing
       
End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub


Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8910 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 4000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 3500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 250
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 850
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 250
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9615 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 850
        scrHScroll.Width = Me.Width - 250
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub imgIcon_Click()

If txtYear = "" Then
    MsgBox "You have to enter a year!"
    txtYear.SetFocus
    Exit Sub
Else
    If Val(txtYear) > Year(Date) + 100 Or Val(txtYear) < Year(Date) - 100 Then
        MsgBox "Year is not a valid year"
        txtYear.SetFocus
        Exit Sub
    End If
End If
frmPayPeriodList.SelectedYear = Val(txtYear)
frmPayPeriodList.Show 1
txtWeek = glbWeek
dlpDateRange(0) = glbFrom
dlpDateRange(1) = glbTo
End Sub

Private Sub medDOW_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
If ReportSel = "PerDayRpt" Then
    If optGrouping(1) Then
        comGroup(2).Enabled = False
        comGroup(2).ListIndex = 1
        comGroup(1).ListIndex = 2
    Else
        comGroup(2).Enabled = True
        comGroup(2).ListIndex = 0
        comGroup(1).ListIndex = 0
    End If
Else
    If optGrouping(1) Or optGrouping(3) Then
        comGroup(2).Enabled = False
        comGroup(2).ListIndex = 1
        comGroup(1).ListIndex = 0 '2
        chkComments.Visible = False
    
        'Ticket #24655 - WDGPHU
        lblYear.Visible = False
        txtYear.Visible = False
        lblWeek.Visible = False
        txtWeek.Visible = False
        imgIcon.Visible = False
        
        chkSup.Enabled = True
        chkUseHireDate.Enabled = True
        chkCostHis.Enabled = True
        chkInclHIS.Enabled = True
        
        If optGrouping(3) And (frmRAttend.Caption = "Attendance Report" Or frmRAttend.Caption = "Attendance Report Including Historical Data") Then
            comGroup(1).ListIndex = 5
        End If
        
    ElseIf optGrouping(2) Then    'Ticket #24655 - WDGPHU
        comGroup(2).Enabled = False
        comGroup(2).ListIndex = 1
        comGroup(1).ListIndex = 0 '2
        'chkComments.Visible = False
        chkComments.Visible = True  'Ticket #27892 - They want to have option to see the Comments in the For Payroll Report.
        chkComments.Left = 9120
        chkComments.Top = 7125
        
        lblYear.Visible = True
        txtYear.Visible = True
        lblWeek.Visible = True
        txtWeek.Visible = True
        imgIcon.Visible = True
        
        chkSup.Enabled = False
        chkUseHireDate.Enabled = False
        chkCostHis.Enabled = False
        chkInclHIS.Enabled = False
    Else
        comGroup(2).Enabled = True
        comGroup(2).ListIndex = 0
        comGroup(1).ListIndex = 0
        chkComments.Visible = True
    
        'Ticket #24655 - WDGPHU
        lblYear.Visible = False
        txtYear.Visible = False
        lblWeek.Visible = False
        txtWeek.Visible = False
        imgIcon.Visible = False
    
        chkSup.Enabled = True
        chkUseHireDate.Enabled = True
        chkCostHis.Enabled = True
        chkInclHIS.Enabled = True
    End If
End If

End Sub

Private Sub SELATTWRK()
Dim SQLQ As String
Dim USQLQ As String
Dim xFieldList As String

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

'Set CN001 = New ADODB.Connection
'CN001.CommandTimeout = 600
'CN001.Open glbAdoIHRDB

'gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
DoEvents
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
DoEvents
'If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
If Not glbOracle Then
    If glbMulti Then
        'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
        If glbCompSerial = "S/N - 2241W" Then   'Granite Club
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            If ReportSel = "COS" Then     'Ticket #12962
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            Else
                SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            End If
        Else
            If ReportSel = "COS" Then
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB  "
            Else
                'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
                'the Job Code in Attendace.
                USQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
            End If
        End If
    Else
        'SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
    End If
Else
    If ReportSel = "COS" And chkCostHis = vbChecked Then     'Ticket #12962
        If glbMulti Then
            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
        Else
            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        End If
    Else
        If glbMulti Then
            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            USQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE" ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        End If
    End If
End If

If Len(HisSQL) > 1 Then
    If glbOracle Then
        'SQLQ = SQLQ & " AND (" & HisSQL & ")"
        SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
        
        'Hemu
        If glbMulti And ReportSel <> "COS" Then
            SQLQ = SQLQ & " AND (AD_JOB IS NULL OR AD_JOB ='')"
            
            USQLQ = USQLQ & " AND (" & HisSQL & ")"
            USQLQ = USQLQ & " AND (AD_JOB IS NOT NULL OR AD_JOB <>'')"
        End If
    Else
        If ReportSel = "COS" Then 'And chkCostHis = vbChecked Then     'Ticket #12962
            'Ticket #13082 - Begin
            'SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            If glbMulti Then
                SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            Else
                'Ticket #19339
                'SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
                SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            End If
            'Ticket #13082 - End
        Else
            'SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
            SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            
            'Hemu
            If glbMulti Then
                If glbCompSerial <> "S/N - 2241W" Then
                    SQLQ = SQLQ & " AND (AD_JOB IS NULL OR AD_JOB ='')"
                    
                    'USQLQ = USQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
                    USQLQ = USQLQ & " WHERE (" & HisSQL & ")"
                    USQLQ = USQLQ & " AND (AD_JOB IS NOT NULL OR AD_JOB <>'')"
                End If
            End If
        End If
    End If
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
DoEvents

'Hemu
If glbMulti And ReportSel <> "COS" And glbCompSerial <> "S/N - 2241W" Then
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute USQLQ
    gdbAdoIhr001.CommitTrans
    DoEvents
End If

If ReportSel = "HIS" Or chkInclHIS Then
    MDIMain.panHelp(0).FloodPercent = 60
    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
    If Not glbOracle Then
        If glbMulti Then
            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            USQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        End If
    Else
        If glbMulti Then
            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            USQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        End If
    End If
    If Len(HisSQL) > 1 Then
        If glbOracle Then
            'SQLQ = SQLQ & "AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            SQLQ = SQLQ & " WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
            
            'Hemu
            If glbMulti And ReportSel <> "COS" Then
                SQLQ = SQLQ & " AND (AH_JOB IS NULL OR AH_JOB ='')"
                
                USQLQ = USQLQ & "AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
                USQLQ = USQLQ & " AND (AH_JOB IS NOT NULL OR AH_JOB <>'')"
            End If
        Else
            'SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            SQLQ = SQLQ & "WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
            
            'Hemu
            If glbMulti And ReportSel <> "COS" Then
                SQLQ = SQLQ & " AND (AH_JOB IS NULL OR AH_JOB ='')"
                
                'USQLQ = USQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
                USQLQ = USQLQ & "WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
                USQLQ = USQLQ & " AND (AH_JOB IS NOT NULL OR AH_JOB <>'')"
            End If
        End If
    End If
    MDIMain.panHelp(0).FloodPercent = 65
    If glbCompSerial <> "S/N - 2241W" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    DoEvents
    
    'Hemu
    If glbMulti And ReportSel <> "COS" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute USQLQ
        gdbAdoIhr001.CommitTrans
        DoEvents
    End If

End If

If chkCostHis = vbChecked And Not glbMulti And Not glbCompSerial = "S/N - 2192W" Then   'Not County of Essex - they are not multi-pos and they change rate in Timesheet
    Dim rsIn As New ADODB.Recordset
    Dim rsOUT As New ADODB.Recordset
    
    If Not (glbOracle Or glbSQL) Then Pause (5)
    
    SQLQ = "SELECT AD_ATT_ID, AD_SALARY, AD_SALCD, AD_WHRS, AD_EMPNBR, AD_DOA FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
    rsOUT.Open SQLQ, gdbAdoIhr001W, adOpenForwardOnly, adLockOptimistic, adCmdText
   
    If rsOUT.EOF = False And rsOUT.BOF = False Then
        Do '
            'Historical Salary
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_TOTAL, SH_SALCD From HR_SALARY_HISTORY WHERE (SH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (SH_EDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY SH_EDATE DESC"
            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            
            If rsIn.EOF = False And rsIn.BOF = False Then
                rsIn.MoveFirst
                If glbCompSerial = "S/N - 2373W" Then
                    rsOUT("AD_SALARY") = rsIn("SH_TOTAL")
                Else
                    rsOUT("AD_SALARY") = rsIn("SH_SALARY")
                End If
                rsOUT("AD_SALCD") = rsIn("SH_SALCD")
            Else
                rsOUT("AD_SALARY") = 0
            End If
            rsIn.Close
            'Historical Position Weekly hours
            SQLQ = "SELECT JH_SDATE, JH_WHRS From HR_JOB_HISTORY WHERE (JH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (JH_SDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY JH_SDATE DESC"
            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsIn.EOF = False And rsIn.BOF = False Then
                rsIn.MoveFirst
                rsOUT("AD_WHRS") = rsIn("JH_WHRS")
            Else
                rsOUT("AD_WHRS") = 0
            End If
            rsIn.Close
            rsOUT.Update
            
            'If Not (glbOracle Or glbSQL) Then Pause (5)

            rsOUT.MoveNext
        Loop Until rsOUT.EOF
    End If
    rsOUT.Close
    DoEvents
    Set rsIn = Nothing
    Set rsOUT = Nothing

'Ticket #17232
'ElseIf glbCompSerial = "S/N - 2241W" And glbMulti And ReportSel = "COS" Then   'Granite Club
'    'Ticket #16017 - Calculate Stat Worked @1.5
'    SQLQ = "UPDATE HRATTWRK SET AD_HRS = (AD_HRS * 1.5) WHERE AD_REASON = 'STAT' AND AD_WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute SQLQ
'    gdbAdoIhr001.CommitTrans
'    DoEvents
End If

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

Private Sub SELATTWRK1()
Dim SQLQ As String
Dim USQLQ As String
Dim xFieldList As String

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

'Set CN001 = New ADODB.Connection
'CN001.CommandTimeout = 600
'CN001.Open glbAdoIHRDB

'gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
DoEvents
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
DoEvents
'If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
If Not glbOracle Then
    If glbMulti Then
        'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
        If glbCompSerial = "S/N - 2241W" Then   'Granite Club
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            If ReportSel = "COS" Then     'Ticket #12962
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            Else
                'Ticket #20039
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            End If
        Else
            If ReportSel = "COS" Then
                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB  "
            Else
                'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
                'the Job Code in Attendace.
                'Ticket #20039
                'USQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
                If chkSup.Value = vbChecked Then    'Ticket #21731
                    SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
                Else
                    SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
                End If
            End If
        End If
    Else
        'SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        If chkSup.Value = vbChecked Then    'Ticket #21731
            SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        Else
            SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        End If
    End If
Else
    If ReportSel = "COS" And chkCostHis = vbChecked Then     'Ticket #12962
        If glbMulti Then
            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
        Else
            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        End If
    Else
        If glbMulti Then
            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            'Ticket #20039
            'USQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE" ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        End If
    End If
End If

'Ticket #21631
If (chkUseHireDate.Visible = True And chkUseHireDate.Value = vbChecked) And (ReportSel <> "COS" And chkCostHis <> vbChecked) Then
    SQLQ = SQLQ & ", HREMP "
End If

If Len(HisSQL) > 1 Then
    If glbOracle Then
        'SQLQ = SQLQ & " AND (" & HisSQL & ")"
        SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
        
        'Hemu
        If glbMulti And ReportSel <> "COS" Then
            'Ticket #20039
            'SQLQ = SQLQ & " AND (AD_JOB IS NULL OR AD_JOB ='')"
            
            'Ticket #20039
            'USQLQ = USQLQ & " AND (" & HisSQL & ")"
            'USQLQ = USQLQ & " AND (AD_JOB IS NOT NULL OR AD_JOB <>'')"
        End If
    Else
        If ReportSel = "COS" Then 'And chkCostHis = vbChecked Then     'Ticket #12962
            'Ticket #13082 - Begin
            'SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            If glbMulti Then
                SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            Else
                'Ticket #19339
                'SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
                SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            End If
            'Ticket #13082 - End
        Else
            'SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
            SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
            
            'Hemu
            If glbMulti Then
                'Ticket #20039
                'If glbCompSerial <> "S/N - 2241W" Then
                '    SQLQ = SQLQ & " AND (AD_JOB IS NULL OR AD_JOB ='')"
                
                    ''USQLQ = USQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
                '    USQLQ = USQLQ & " WHERE (" & HisSQL & ")"
                '    USQLQ = USQLQ & " AND (AD_JOB IS NOT NULL OR AD_JOB <>'')"
                'End If
            End If
        End If
    End If
    
    'Ticket #21631
    If (chkUseHireDate.Visible = True And chkUseHireDate.Value = vbChecked) And (ReportSel <> "COS" And chkCostHis <> vbChecked) Then
        SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR "
    End If
    
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
DoEvents

'Hemu
'Ticket #20039
'If glbMulti And ReportSel <> "COS" And glbCompSerial <> "S/N - 2241W" Then
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute USQLQ
'    gdbAdoIhr001.CommitTrans
'    DoEvents
'End If

If ReportSel = "HIS" Or chkInclHIS Then
    MDIMain.panHelp(0).FloodPercent = 60
    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
    If Not glbOracle Then
        If glbMulti Then
            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            'Ticket #20039
            'USQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
            If chkSup.Value = vbChecked Then    'Ticket #21731
                SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
            Else
                SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
            End If
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
            If chkSup.Value = vbChecked Then    'Ticket #21731
                SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
            Else
                SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
            End If
        End If
    Else
        If glbMulti Then
            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
            'the Job Code in Attendace.
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            'Ticket #20039
            'USQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        Else
            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
        End If
    End If
    
    'Ticket #21631
    If (chkUseHireDate.Visible = True And chkUseHireDate.Value = vbChecked) And (ReportSel <> "COS" And chkCostHis <> vbChecked) Then
        SQLQ = SQLQ & ", HREMP "
    End If
    
    If Len(HisSQL) > 1 Then
        If glbOracle Then
            'SQLQ = SQLQ & "AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            SQLQ = SQLQ & " WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
            
            'Hemu
            'Ticket #20039
            'If glbMulti And ReportSel <> "COS" Then
            '    SQLQ = SQLQ & " AND (AH_JOB IS NULL OR AH_JOB ='')"
            '
            '    USQLQ = USQLQ & "AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            '    USQLQ = USQLQ & " AND (AH_JOB IS NOT NULL OR AH_JOB <>'')"
            'End If
        Else
            'SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            SQLQ = SQLQ & "WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
            
            'Hemu
            'Ticket #20039
            'If glbMulti And ReportSel <> "COS" Then
            '    SQLQ = SQLQ & " AND (AH_JOB IS NULL OR AH_JOB ='')"
                
                ''USQLQ = USQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
            '    USQLQ = USQLQ & "WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
            '    USQLQ = USQLQ & " AND (AH_JOB IS NOT NULL OR AH_JOB <>'')"
            'End If
        End If
        
        'Ticket #21631
        If (chkUseHireDate.Visible = True And chkUseHireDate.Value = vbChecked) And (ReportSel <> "COS" And chkCostHis <> vbChecked) Then
            SQLQ = SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_EMPNBR = HREMP.ED_EMPNBR "
        End If
        
    End If
    MDIMain.panHelp(0).FloodPercent = 65
    'Ticket #20039
    'If glbCompSerial <> "S/N - 2241W" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    'End If
    DoEvents
    
    'Hemu
    'Ticket #20039
    'If glbMulti And ReportSel <> "COS" Then
    '    gdbAdoIhr001.BeginTrans
    '    gdbAdoIhr001.Execute USQLQ
    '    gdbAdoIhr001.CommitTrans
    '    DoEvents
    'End If

End If

'Ticket #27095 - Opening up for ALL Attendance reports
'Ticket #26272 Franks 01/08/2015
'If ReportSel = "HIS" And chkInclTerm.Value Then
If chkInclTerm.Value Then
    MDIMain.panHelp(0).FloodPercent = 80
    'update AD_KEY_EMPNBR for active employees
    SQLQ = "UPDATE HRATTWRK SET AD_KEY_EMPNBR = '1_' + LTRIM(RTRIM(STR(AD_EMPNBR)) ) "
    SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    'SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP,AD_KEY_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
    SQLQ = SQLQ & ",'0_' + LTRIM(RTRIM(STR(Term_ATTENDANCE.AD_EMPNBR)) ) +  LTRIM(RTRIM(STR(Term_ATTENDANCE.TERM_SEQ)) ) "
    
    If glbMulti Then
        'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
        'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
        'the Job Code in Attendace.
        'Ticket #20039
        'USQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
        If chkSup.Value = vbChecked Then    'Ticket #21731
            SQLQ = SQLQ & " FROM Term_ATTENDANCE LEFT OUTER JOIN Term_JOB_HISTORY ON (Term_ATTENDANCE.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
        Else
            SQLQ = SQLQ & " FROM Term_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
        End If
    Else
        'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        If chkSup.Value = vbChecked Then    'Ticket #21731
            SQLQ = SQLQ & " FROM Term_ATTENDANCE LEFT OUTER JOIN Term_JOB_HISTORY ON (Term_ATTENDANCE.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ) "
        Else
            SQLQ = SQLQ & " FROM Term_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        End If
    End If

    
    ''Ticket #21631
    'If (chkUseHireDate.Visible = True And chkUseHireDate.Value = vbChecked) And (ReportSel <> "COS" And chkCostHis <> vbChecked) Then
        SQLQ = SQLQ & ",Term_HREMP,TERM_HRTRMEMP "
    'End If
    
    If Len(HisSQL) > 1 Then
        SQLQ = SQLQ & "WHERE (" & Replace(HisSQL, "HREMP", "Term_HREMP") & ") "
        SQLQ = SQLQ & "AND Term_ATTENDANCE.TERM_SEQ = Term_HREMP.TERM_SEQ "
        SQLQ = SQLQ & "AND TERM_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ  "
        SQLQ = SQLQ & "AND TERM_HRTRMEMP.TERM_DOR IS NULL "
    End If
    MDIMain.panHelp(0).FloodPercent = 85
    'Ticket #20039
    'If glbCompSerial <> "S/N - 2241W" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    'End If
        
    DoEvents

End If

If chkCostHis = vbChecked And Not glbMulti And Not glbCompSerial = "S/N - 2192W" Then   'Not County of Essex - they are not multi-pos and they change rate in Timesheet
    Dim rsIn As New ADODB.Recordset
    Dim rsOUT As New ADODB.Recordset
    
    If Not (glbOracle Or glbSQL) Then Pause (5)
    
    SQLQ = "SELECT AD_ATT_ID, AD_SALARY, AD_SALCD, AD_WHRS, AD_EMPNBR, AD_DOA FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
    rsOUT.Open SQLQ, gdbAdoIhr001W, adOpenForwardOnly, adLockOptimistic, adCmdText
   
    If rsOUT.EOF = False And rsOUT.BOF = False Then
        Do '
            'Historical Salary
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_TOTAL, SH_SALCD From HR_SALARY_HISTORY WHERE (SH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (SH_EDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY SH_EDATE DESC"
            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            
            If rsIn.EOF = False And rsIn.BOF = False Then
                rsIn.MoveFirst
                If glbCompSerial = "S/N - 2373W" Then
                    rsOUT("AD_SALARY") = rsIn("SH_TOTAL")
                Else
                    rsOUT("AD_SALARY") = rsIn("SH_SALARY")
                End If
                rsOUT("AD_SALCD") = rsIn("SH_SALCD")
            Else
                rsOUT("AD_SALARY") = 0
            End If
            rsIn.Close
            'Historical Position Weekly hours
            SQLQ = "SELECT JH_SDATE, JH_WHRS From HR_JOB_HISTORY WHERE (JH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (JH_SDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY JH_SDATE DESC"
            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rsIn.EOF = False And rsIn.BOF = False Then
                rsIn.MoveFirst
                rsOUT("AD_WHRS") = rsIn("JH_WHRS")
            Else
                rsOUT("AD_WHRS") = 0
            End If
            rsIn.Close
            rsOUT.Update
            
            'If Not (glbOracle Or glbSQL) Then Pause (5)

            rsOUT.MoveNext
        Loop Until rsOUT.EOF
    End If
    rsOUT.Close
    DoEvents
    Set rsIn = Nothing
    Set rsOUT = Nothing
'Ticket #17232
'ElseIf glbCompSerial = "S/N - 2241W" And glbMulti And ReportSel = "COS" Then   'Granite Club
'    'Ticket #16017 - Calculate Stat Worked @1.5
'    SQLQ = "UPDATE HRATTWRK SET AD_HRS = (AD_HRS * 1.5) WHERE AD_REASON = 'STAT' AND AD_WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute SQLQ
'    gdbAdoIhr001.CommitTrans
'    DoEvents
End If

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub


'Private Sub SELATTWRK()
'Dim SQLQ As String
'Dim xFieldList As String
'
'On Error GoTo AttWrkError
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(1).Caption = " Please Wait"
'MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(0).FloodPercent = 1
'MDIMain.panHelp(0).FloodPercent = 3
'
''Set CN001 = New ADODB.Connection
''CN001.CommandTimeout = 600
''CN001.Open glbAdoIHRDB
'
''gdbAdoIhr001.CommandTimeout = 600
'Screen.MousePointer = HOURGLASS
'MDIMain.panHelp(0).FloodPercent = 15
'DoEvents
'gdbAdoIhr001W.BeginTrans
'gdbAdoIhr001W.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
'gdbAdoIhr001W.CommitTrans
'DoEvents
''If Not glbSQL And Not glbOracle Then Call Pause(1)
'MDIMain.panHelp(0).FloodPercent = 30
'
'xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
'SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
'SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
'If Not glbOracle Then
'    If glbMulti Then
'        'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
'        If glbCompSerial = "S/N - 2241W" Then   'Granite Club
'            'SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'            If ReportSel = "COS" Then     'Ticket #12962
'                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'            Else
'                SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'            End If
'        Else
'            If ReportSel = "COS" Then
'                SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB  "
'            Else
'                'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
'                'the Job Code in Attendace.
'                SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
'            End If
'        End If
'    Else
'        SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
'    End If
'Else
'    If ReportSel = "COS" And chkCostHis = vbChecked Then     'Ticket #12962
'        If glbMulti Then
'            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
'        End If
'    Else
'        If glbMulti Then
'            'Bryan removed " AND HR_ATTENDANCE.AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
'            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
'            'the Job Code in Attendace.
'            'SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        End If
'    End If
'End If
'
'If Len(HisSQL) > 1 Then
'    If glbOracle Then
'        SQLQ = SQLQ & " AND (" & HisSQL & ")"
'    Else
'        If ReportSel = "COS" Then 'And chkCostHis = vbChecked Then     'Ticket #12962
'            'Ticket #13082 - Begin
'            'SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
'            If glbMulti Then
'                SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
'            Else
'                SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
'            End If
'            'Ticket #13082 - End
'        Else
'            SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & HisSQL & ")"
'        End If
'    End If
'End If
'
'MDIMain.panHelp(0).FloodPercent = 45
'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute SQLQ
'gdbAdoIhr001.CommitTrans
'DoEvents
'If ReportSel = "HIS" Or chkInclHIS Then
'    MDIMain.panHelp(0).FloodPercent = 60
'    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
'    SQLQ = SQLQ & in_SQL(glbIHRDBW)
'    SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
'    If Not glbOracle Then
'        If glbMulti Then
'            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
'            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
'            'the Job Code in Attendace.
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
'        End If
'    Else
'        If glbMulti Then
'            'Bryan removed "  AND HR_ATTENDANCE_HISTORY.AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO" Ticket#12764, Jerry suggested that this should have been serial number controlled.
'            'Ticket #14538 - Jerry said to remove the condition to check if the Job Code in Job_History matches
'            'the Job Code in Attendace.
'            'SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        End If
'    End If
'    If Len(HisSQL) > 1 Then
'        If glbOracle Then
'            SQLQ = SQLQ & "AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
'        Else
'            SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & Replace(HisSQL, "AD_", "AH_") & ")"
'        End If
'    End If
'    MDIMain.panHelp(0).FloodPercent = 65
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute SQLQ
'    gdbAdoIhr001.CommitTrans
'    DoEvents
'End If
'
'If chkCostHis = vbChecked And Not glbMulti And Not glbCompSerial = "S/N - 2192W" Then   'Not County of Essex - they are not multi-pos and they change rate in Timesheet
'    Dim rsIn As New ADODB.Recordset
'    Dim rsOUT As New ADODB.Recordset
'
'    If Not (glbOracle Or glbSQL) Then Pause (5)
'
'    SQLQ = "SELECT AD_ATT_ID, AD_SALARY, AD_SALCD, AD_WHRS, AD_EMPNBR, AD_DOA FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
'    rsOUT.Open SQLQ, gdbAdoIhr001W, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If rsOUT.EOF = False And rsOUT.BOF = False Then
'        Do '
'            'Historical Salary
'            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_TOTAL, SH_SALCD From HR_SALARY_HISTORY WHERE (SH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (SH_EDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY SH_EDATE DESC"
'            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'
'            If rsIn.EOF = False And rsIn.BOF = False Then
'                rsIn.MoveFirst
'                If glbCompSerial = "S/N - 2373W" Then
'                    rsOUT("AD_SALARY") = rsIn("SH_TOTAL")
'                Else
'                    rsOUT("AD_SALARY") = rsIn("SH_SALARY")
'                End If
'                rsOUT("AD_SALCD") = rsIn("SH_SALCD")
'            Else
'                rsOUT("AD_SALARY") = 0
'            End If
'            rsIn.Close
'            'Historical Position Weekly hours
'            SQLQ = "SELECT JH_SDATE, JH_WHRS From HR_JOB_HISTORY WHERE (JH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (JH_SDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY JH_SDATE DESC"
'            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'            If rsIn.EOF = False And rsIn.BOF = False Then
'                rsIn.MoveFirst
'                rsOUT("AD_WHRS") = rsIn("JH_WHRS")
'            Else
'                rsOUT("AD_WHRS") = 0
'            End If
'            rsIn.Close
'            rsOUT.Update
'
'            'If Not (glbOracle Or glbSQL) Then Pause (5)
'
'            rsOUT.MoveNext
'        Loop Until rsOUT.EOF
'    End If
'    rsOUT.Close
'    DoEvents
'    Set rsIn = Nothing
'    Set rsOUT = Nothing
''Ticket #17232
''ElseIf glbCompSerial = "S/N - 2241W" And glbMulti And ReportSel = "COS" Then   'Granite Club
''    'Ticket #16017 - Calculate Stat Worked @1.5
''    SQLQ = "UPDATE HRATTWRK SET AD_HRS = (AD_HRS * 1.5) WHERE AD_REASON = 'STAT' AND AD_WRKEMP='" & glbUserID & "'"
''    gdbAdoIhr001.BeginTrans
''    gdbAdoIhr001.Execute SQLQ
''    gdbAdoIhr001.CommitTrans
''    DoEvents
'End If
'
'MDIMain.panHelp(0).FloodPercent = 100
'MDIMain.panHelp(0).FloodType = 0
'MDIMain.panHelp(1).Caption = ""
'MDIMain.panHelp(2).Caption = ""
'Exit Sub
'
'AttWrkError:
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = ""
'    MDIMain.panHelp(2).Caption = ""
'    'If Err.Number = 2147217871 Then MsgBox Err.Description
'    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
'    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
'    Exit Sub
'End Sub


Private Sub scrControl_Change()
    scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 20 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub TxtCases_DblClick()
TxtCases.Text = ""
frmCaseNolp.Show 1
TxtCases.Text = glbCaseFiles

End Sub

Private Sub txtChargeCode_GotFocus()
   Call SetPanHelp(Me.ActiveControl)     'laura 03/03/98
End Sub

Private Sub txtHours_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncentive_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncentive_KeyUp(KeyCode As Integer, Shift As Integer)
    txtIncentive.Text = UCase(txtIncentive.Text)
End Sub

Private Sub txtIncident_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncident_KeyUp(KeyCode As Integer, Shift As Integer)
    txtIncident.Text = UCase(txtIncident.Text)
End Sub

Private Sub txtSeniority_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
 MDIMain.panHelp(0).Caption = Replace(MDIMain.panHelp(0).Caption, "s", "S")
End Sub

Private Sub txtSeniority_KeyUp(KeyCode As Integer, Shift As Integer)
    txtSeniority.Text = UCase(txtSeniority.Text)
End Sub

Private Sub txtShift_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_Job()
Dim EECri As String

If Len(clpJOB.Text) > 0 Then
    HisSQL = HisSQL & " AND (AD_JOB = '" & clpJOB.Text & "') "
    EECri = "{HR_ATTENDANCE.AD_JOB} = '" & clpJOB.Text & "' "
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

Private Sub Update_Att_Salary()
    Dim rsSal As New ADODB.Recordset
    Dim rsATT As New ADODB.Recordset
    Dim xEmpnbr, xFDate, xTDate ', xSalary, xSalcd
    Dim SQLQ
    Dim I
    I = 0
    
    SQLQ = "SELECT SH_EDATE,SH_EMPNBR, SH_SALARY as SH_Sal, SH_SALCD,SH_JOB,SH_WHRS "
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE SH_EMPNBR IN  "
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "')"
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR, SH_EDATE desc "
    
    If glbCompSerial = "S/N - 2373W" Then
        SQLQ = Replace(SQLQ, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
    End If
    
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do Until rsSal.EOF
        If xEmpnbr <> rsSal!SH_EMPNBR Then xTDate = Null Else xTDate = xFDate
        xEmpnbr = rsSal!SH_EMPNBR
        xFDate = rsSal!SH_EDATE
        
        SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_SALARY, AD_SALCD,AD_JOB,AD_WHRS,AD_ATT_ID "

            SQLQ = SQLQ & " FROM HRATTWRK "
            SQLQ = SQLQ & " WHERE AD_EMPNBR=" & xEmpnbr
               
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(xFDate)
        If Not IsNull(xTDate) Then SQLQ = SQLQ & " AND AD_DOA<" & Date_SQL(xTDate)
        'redundant, the data in HRATTWRK already has been limited by HisSQL. Bryan 30/May/2006 Ticket#10902
        'SQLQ = SQLQ & " AND " & Replace(HisSQL, "HREMP ", "HREMP " & in_SQL(glbIHRDB) & " ")
        SQLQ = SQLQ & " AND AD_WRKEMP='" & glbUserID & "'"
        
        If IsDate(xFDate) Then
            rsATT.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            gdbAdoIhr001W.BeginTrans
            Do Until rsATT.EOF
                rsATT!AD_SALARY = rsSal!SH_SAL
                rsATT!AD_SALCD = rsSal!SH_SALCD
                rsATT!AD_JOB = rsSal!SH_JOB
                rsATT!AD_WHRS = rsSal!SH_WHRS
                rsATT.Update
                rsATT.MoveNext
                I = I + 1
                MDIMain.panHelp(0).Caption = I & " Records Processed"
            Loop
            gdbAdoIhr001W.CommitTrans
            rsATT.Close
        End If
        rsSal.MoveNext
    Loop
    rsSal.Close
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

Private Sub ScreenPerDayRpt()
lblEENum(1).Visible = False
lblChargeCode.Visible = False
lblDOW(0).Visible = False
lblDOW(1).Visible = False
lblShift.Visible = False
lblIncident.Visible = False
lblRegion.Visible = False
lblAdmin.Visible = False
lblAccount.Visible = False
lblIncentive.Visible = False
Label1(5).Visible = False
lblJOB.Visible = False
lblPosGroup.Visible = False
lblHours.Visible = False
lblHoursLess.Visible = False
lblSen1.Visible = False
lblSection.Visible = False
lblMachine.Visible = False

clpChrgCode.Visible = False
medDOW(0).Visible = False
medDOW(1).Visible = False
txtShift.Visible = False
txtIncident.Visible = False
clpCode(9).Visible = False
clpCode(10).Visible = False
clpCode(11).Visible = False
txtIncentive.Visible = False
elpSUP(1).Visible = False
txtShift.Visible = False
cmbUpload.Visible = False
clpJOB.Visible = False
clpCode(7).Visible = False
txtHours.Visible = False
txtHoursLess.Visible = False
txtSeniority.Visible = False
clpCode(5).Visible = False
clpCode(8).Visible = False
chkComments.Visible = False
chkAbsence.Visible = False
chkSup.Visible = False
lblGrp(2).Visible = False
lblGrp(3).Visible = False
comGroup(2).Visible = False
comGroup(3).Visible = False
comGroup(0).Text = "(none)"
comGroup(0).Enabled = False
lblFromTo.Top = lblAttCodes.Top
lblAttCodes.Top = lblEENum(1).Top
dlpDateRange(0).Top = clpCode(2).Top
dlpDateRange(1).Top = clpCode(2).Top
clpCode(2).Top = elpSUP(1).Top

End Sub

Private Sub GetDisData()
Dim I As Integer
Dim listCourseTypes
Dim listCourseCodes
Dim listJobs
    For I = 1 To 10
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
    If Len(clpCode(6).Text) > 0 Then
        fDis(5) = "Dis(5) = '" & lStr(lblStatus) & " = ' + chr(34) + '" & Replace(clpCode(6).Caption, "'", "''") & "' + chr(34)"
    End If
    If Len(clpPT.Text) > 0 Then
        fDis(6) = "Dis(6) = '" & lStr(lblPT) & " = ' + chr(34) + '" & Replace(clpPT.Caption, "'", "''") & "' + chr(34)"
    End If

End Sub

Private Sub Wellington_Terrace_Attendance_Report()
Dim SQLQ As String
Dim rsAttend As New ADODB.Recordset
Dim rsComments As New ADODB.Recordset
Dim xHrsDay As Double
Dim xComments As String
Dim xDiv As String
Dim xPrvEmpTot, xPrvReasonTot As Double
Dim xPrvReasonHTot, xPrvReasonDTot, xPrvEmpHTot, xPrvEmpDTot     As Double
Dim xPrvEmp, xPrvReason As String
Dim xReasonFlg, xEmpFlg As Boolean

Dim exApp As Excel.Application
Dim exBook As Excel.Workbook
Dim exSheet As Excel.Worksheet
Dim xlsFileTmp As String
Dim xlsFileMat As String
Dim I, totNum
Dim xRow As Integer
Dim xExcelRptPath As String

On Error GoTo EnvironErr

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

Screen.MousePointer = HOURGLASS

'Ticket #22034 - Get Excel reports path
If gsTRAININGMATRIX Then
    xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
End If
If Len(xExcelRptPath) = 0 Then
    xExcelRptPath = glbIHRREPORTS
End If

MDIMain.panHelp(0).FloodPercent = 5
DoEvents


'sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")

'Retrieve the Attendance records matching the selection criteria and for the respective Attendance
'retrieve the Comments Type AMT with matching Attendance Date.
SQLQ = "SELECT AD_EMPNBR, AD_DOA, AD_HRS, AD_REASON, AD_COMM, ED_SURNAME, ED_FNAME, ED_DIV FROM HR_ATTENDANCE "
SQLQ = SQLQ & ", HREMP "
SQLQ = SQLQ & " WHERE HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR "
SQLQ = SQLQ & " AND (" & HisSQL & ")"
'SQLQ = SQLQ & " AND ED_SECTION = 'TERR'"

'Ticket #29704 - Include Attendance History
SQLQ = SQLQ & " UNION "

SQLQ = SQLQ & "SELECT AH_EMPNBR AS AD_EMPNBR, AH_DOA AS AD_DOA, AH_HRS AS AD_HRS, AH_REASON AS AD_REASON, AH_COMM AS AD_COMM, ED_SURNAME, ED_FNAME, ED_DIV FROM HR_ATTENDANCE_HISTORY "
SQLQ = SQLQ & ", HREMP "
SQLQ = SQLQ & " WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HREMP.ED_EMPNBR "
SQLQ = SQLQ & " AND (" & Replace(Replace(HisSQL, "HR_ATTENDANCE.", "HR_ATTENDANCE_HISTORY."), "AD_", "AH_") & ")"

SQLQ = SQLQ & " ORDER BY ED_DIV, ED_SURNAME, ED_FNAME, AD_REASON, AD_DOA"
rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsAttend.EOF Then
    totNum = rsAttend.RecordCount: I = 0

    'Initialise/Open Excel Report file
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2262_AttendRptTmp.xls"
    
    'Ticket #22034 - May need to save the report in different path
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttendRpt_" & Trim(glbUserID) & ".xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "AttendRpt_" & Trim(glbUserID) & ".xls"
    

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


    'Print the report headers
    exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy hh:mm")
    
    exSheet.Cells(2, 1) = "Wellington Terrace Attendance Report"
    exSheet.Cells(3, 1) = Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " - " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")

    'Row 5 column headers
    'Employee #(1), Employee Name(2), Attendance Reason(3), Notes(4), Date(5), Day of Week(6), Hours(7),
    'Days(8), Comments(9)
    
    'Initialise
    xRow = 5
    xDiv = ""
    xComments = ""
    xPrvReason = ""
    xPrvEmp = ""
    xPrvReasonHTot = 0
    xPrvReasonDTot = 0
    xPrvEmpHTot = 0
    xPrvEmpDTot = 0
    
    Do While Not rsAttend.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        
        xComments = ""
        SQLQ = "SELECT * FROM HR_COMMENTS "
        SQLQ = SQLQ & " WHERE CO_EMPNBR = " & rsAttend("AD_EMPNBR")
        SQLQ = SQLQ & " AND CO_EDATE = " & Date_SQL(rsAttend("AD_DOA"))
        SQLQ = SQLQ & " AND CO_TYPE = 'AMT'"
        rsComments.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If Not rsComments.EOF Then
            xComments = IIf(IsNull(rsComments("CO_COMMENTS")), "", rsComments("CO_COMMENTS"))
        End If
        rsComments.Close
        Set rsComments = Nothing
        
        'Display Division and header columns if beginning of the list
        If xRow = 5 Then
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Cells(xRow, 1) = lStr("Division:")
            If Not IsNull(rsAttend("ED_DIV")) Then
                exSheet.Cells(xRow, 2) = Get_Division_Name(rsAttend("ED_DIV"))
            End If
                        
            'Column Header
            xRow = xRow + 2
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Cells(xRow, 1) = "Employee #"
            exSheet.Cells(xRow, 2) = "Employee Name"
            exSheet.Cells(xRow, 3) = "Code"
            exSheet.Cells(xRow, 4) = "Notes"
            exSheet.Cells(xRow, 5) = "Date"
            exSheet.Cells(xRow, 6) = "Day of Week"
            exSheet.Cells(xRow, 7) = "Hours"
            exSheet.Cells(xRow, 8) = "Days"
            exSheet.Cells(xRow, 9) = "Comments"
            
            'Set the values for future ref.
            If Not IsNull(rsAttend("ED_DIV")) Then
                xDiv = rsAttend("ED_DIV")
            End If
            xPrvReason = rsAttend("AD_REASON")
            xPrvEmp = rsAttend("AD_EMPNBR") 'rsAttend("ED_SURNAME") & ", " & rsAttend("ED_FNAME")
            xReasonFlg = True
            xEmpFlg = True
            
            xRow = xRow + 1
        End If
        
        'Check if the Division has changed. If so then display the sum of hours by Reason and Employee
        'and create a new division section
        If xDiv <> rsAttend("ED_DIV") Then
            
            'Display totals and new division
            If xRow <> 5 Then
                'Reason and Employee Totals
                'xRow = xRow + 1
                exSheet.Rows(xRow).Font.Bold = True
                exSheet.Cells(xRow, 2) = "Total for " & GetTABLDesc("ADRE", xPrvReason)
                exSheet.Cells(xRow, 7) = Format(xPrvReasonHTot, "#.00")
                exSheet.Cells(xRow, 8) = Format(xPrvReasonDTot, "#.00")
                xRow = xRow + 1
                exSheet.Rows(xRow).Font.Bold = True
                exSheet.Cells(xRow, 2) = "Total for " & GetEmpData(xPrvEmp, "ED_SURNAME") & ", " & GetEmpData(xPrvEmp, "ED_FNAME")
                exSheet.Cells(xRow, 7) = Format(xPrvEmpHTot, "#.00")
                exSheet.Cells(xRow, 8) = Format(xPrvEmpDTot, "#.00")
                
                'New Division
                xRow = xRow + 3
                exSheet.Rows(xRow).Font.Bold = True
                exSheet.Rows(xRow).Font.Size = 9
                exSheet.Cells(xRow, 1) = lStr("Division:")
                exSheet.Cells(xRow, 2) = Get_Division_Name(rsAttend("ED_DIV"))
                
                'Column Header
                xRow = xRow + 2
                exSheet.Rows(xRow).Font.Bold = True
                exSheet.Cells(xRow, 1) = "Employee #"
                exSheet.Cells(xRow, 2) = "Employee Name"
                exSheet.Cells(xRow, 3) = "Code"
                exSheet.Cells(xRow, 4) = "Notes"
                exSheet.Cells(xRow, 5) = "Date"
                exSheet.Cells(xRow, 6) = "Day of Week"
                exSheet.Cells(xRow, 7) = "Hours"
                exSheet.Cells(xRow, 8) = "Days"
                exSheet.Cells(xRow, 9) = "Comments"
                
                'Set the values for future ref.
                xDiv = rsAttend("ED_DIV")
                xPrvReason = rsAttend("AD_REASON")
                xPrvEmp = rsAttend("AD_EMPNBR") 'rsAttend("ED_SURNAME") & ", " & rsAttend("ED_FNAME")
                xReasonFlg = True
                xEmpFlg = True
                xPrvReasonHTot = 0
                xPrvReasonDTot = 0
                xPrvEmpHTot = 0
                xPrvEmpDTot = 0
                
                xRow = xRow + 1
            End If
        End If
                
        'Get current hours per day
        xHrsDay = 0
        xHrsDay = GetJHData(rsAttend("AD_EMPNBR"), "JH_DHRS", 0)
                                
        'If Employee is different from Previous, display Reason Total and Employee Total
        If xPrvEmp <> rsAttend("AD_EMPNBR") Then 'rsAttend("ED_SURNAME") & ", " & rsAttend("ED_FNAME") Then
            'Reason and Employee Totals
            'xRow = xRow + 1
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Cells(xRow, 2) = "Total for " & GetTABLDesc("ADRE", xPrvReason)
            exSheet.Cells(xRow, 7) = Format(xPrvReasonHTot, "#.00")
            exSheet.Cells(xRow, 8) = Format(xPrvReasonDTot, "#.00")
            xRow = xRow + 1
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Cells(xRow, 2) = "Total for " & GetEmpData(xPrvEmp, "ED_SURNAME") & ", " & GetEmpData(xPrvEmp, "ED_FNAME")
            exSheet.Cells(xRow, 7) = Format(xPrvEmpHTot, "#.00")
            exSheet.Cells(xRow, 8) = Format(xPrvEmpDTot, "#.00")
            
            xPrvReasonHTot = 0
            xPrvReasonDTot = 0
            xPrvEmpHTot = 0
            xPrvEmpDTot = 0
            xPrvReason = rsAttend("AD_REASON")
            xPrvEmp = rsAttend("AD_EMPNBR") 'rsAttend("ED_SURNAME") & ", " & rsAttend("ED_FNAME")
            xEmpFlg = True
            xReasonFlg = True
            xRow = xRow + 2
        Else
            'xReasonFlg = False
            'xEmpFlg = False
            
            'Sum hours by Reason
            'xPrvReasonHTot = xPrvReasonHTot + IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS"))
            'If xHrsDay <> 0 Then
            '    xPrvReasonDTot = xPrvReasonDTot + (IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay)
            'End If
            
            'xPrvEmpHTot = xPrvEmpHTot + IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS"))
            'If xHrsDay <> 0 Then
            '    xPrvEmpDTot = xPrvEmpDTot + (IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay)
            'End If
        End If
        
        'If Reason is different from previous, display Reason Total
        If xPrvReason <> rsAttend("AD_REASON") Then
            'Reason Total
            'xRow = xRow + 1
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Cells(xRow, 2) = "Total for " & GetTABLDesc("ADRE", xPrvReason)
            exSheet.Cells(xRow, 7) = Format(xPrvReasonHTot, "#.00")
            exSheet.Cells(xRow, 8) = Format(xPrvReasonDTot, "#.00")
            
            xPrvReasonHTot = 0
            xPrvReasonDTot = 0
            xPrvReason = rsAttend("AD_REASON")
            xReasonFlg = True
            xRow = xRow + 1
        Else
            'xReasonFlg = False
            
            'Sum hours by reason
            'xPrvReasonHTot = xPrvReasonHTot + IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS"))
            'If xHrsDay <> 0 Then
            '    xPrvReasonDTot = xPrvReasonDTot + (IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay)
            'End If
        End If
        
        'Print the Attendance record with Comments Type AMT's Comments
        If xEmpFlg = True Or xRow = 8 Then  'Employee changed, display on the report
            exSheet.Cells(xRow, 1) = rsAttend("AD_EMPNBR")
            exSheet.Cells(xRow, 2) = rsAttend("ED_SURNAME") & ", " & rsAttend("ED_FNAME")
            xEmpFlg = False
        End If
        If xReasonFlg = True Or xRow = 8 Then   'Reason changed, display on the report
            exSheet.Cells(xRow, 3) = GetTABLDesc("ADRE", rsAttend("AD_REASON"))
            xReasonFlg = False
        End If
        exSheet.Cells(xRow, 4) = xComments
        exSheet.Cells(xRow, 5) = Format(rsAttend("AD_DOA"), "mm/dd/yyyy")
        exSheet.Cells(xRow, 6) = WeekdayName(Weekday(rsAttend("AD_DOA")))
        exSheet.Cells(xRow, 7) = Format(IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")), "#.00")
        If xHrsDay <> 0 Then
            exSheet.Cells(xRow, 8) = Format(IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay, "#.00")
        End If
        exSheet.Cells(xRow, 9) = rsAttend("AD_COMM")
        
        xPrvReasonHTot = xPrvReasonHTot + IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS"))
        If xHrsDay <> 0 Then
            xPrvReasonDTot = xPrvReasonDTot + (IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay)
        End If
        
        xPrvEmpHTot = xPrvEmpHTot + IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS"))
        If xHrsDay <> 0 Then
            xPrvEmpDTot = xPrvEmpDTot + (IIf(IsNull(rsAttend("AD_HRS")), 0, rsAttend("AD_HRS")) / xHrsDay)
        End If
        
        'Move to next row
        xRow = xRow + 1
        
        rsAttend.MoveNext
    Loop
    
    exSheet.Rows(xRow).Font.Bold = True
    exSheet.Cells(xRow, 2) = "Total for " & GetTABLDesc("ADRE", xPrvReason)
    exSheet.Cells(xRow, 7) = Format(xPrvReasonHTot, "#.00")
    exSheet.Cells(xRow, 8) = Format(xPrvReasonDTot, "#.00")
    xRow = xRow + 1
    exSheet.Rows(xRow).Font.Bold = True
    exSheet.Cells(xRow, 2) = "Total for " & GetEmpData(xPrvEmp, "ED_SURNAME") & ", " & GetEmpData(xPrvEmp, "ED_FNAME")
    exSheet.Cells(xRow, 7) = Format(xPrvEmpHTot, "#.00")
    exSheet.Cells(xRow, 8) = Format(xPrvEmpDTot, "#.00")

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
Else
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
    
    MsgBox "No employees found in this selection criteria."
End If
rsAttend.Close
Set rsAttend = Nothing

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

EnvironErr:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    
    Set exSheet = Nothing
    Set exBook = Nothing
    Set exApp = Nothing
    
End Sub

Private Sub txtWeek_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Private Sub txtWeek_DblClick()
Call imgIcon_Click
End Sub

Private Sub txtWeek_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtWeek_LostFocus()
If txtWeek = "" Then
    dlpDateRange(0) = ""
    dlpDateRange(1) = ""
Else
    'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
End If
End Sub

Function getDateRange(theClientNumber, thePayNbr, theYear)
Dim rsPayPeriod As New ADODB.Recordset
Dim SQLQ, intNum

On Error Resume Next

getDateRange = "|"

If Not IsNumeric(thePayNbr) Then Exit Function
If Not IsNumeric(theYear) Then Exit Function

SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_payperiod WHERE PP_PAYP='" & theClientNumber & "'"
SQLQ = SQLQ & " and PP_NBR = " & thePayNbr
SQLQ = SQLQ & " and PP_YEAR = '" & theYear & "'"
rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

If Not rsPayPeriod.EOF Then
    getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
   
End If
rsPayPeriod.Close
Exit Function

End Function

Private Sub WDGPHU_PayrollAttendance()
Dim rsAttend As New ADODB.Recordset
Dim rsAttWRK As New ADODB.Recordset
Dim SQLQ As String
Dim USQLQ As String
Dim xFieldList As String
Dim xSFieldList As String

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
DoEvents

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
DoEvents

MDIMain.panHelp(0).FloodPercent = 30

xFieldList = "AD_EMPNBR,AD_REASON,AD_CHRGCODE,AD_REGION,AD_HRS,AD_SALARY,AD_SALCD,AD_WCBNBR"
'Ticket #28539 - Only add Comments if Show Comments is checked otherwise the group by is showing an issue.
If chkComments Then
    xFieldList = xFieldList & ",AD_COMM"    'Ticket #27892 - Added Comments
End If
xSFieldList = "AD_EMPNBR,AD_REASON,AD_CHRGCODE,AD_REGION,SUM(AD_HRS) AS AD_HRS,AD_SALARY,AD_SALCD, '" & txtWeek.Text & "' AS AD_WCBNBR"
'Ticket #28539 - Only add Comments if Show Comments is checked otherwise the group by is showing an issue.
If chkComments Then
    xSFieldList = xSFieldList & ",AD_COMM"      'Ticket #27892 - Added Comments
End If
'SQLQ = "SELECT AD_EMPNBR,AD_REASON,AD_CHRGCODE,AD_REGION,SUM(AD_HRS) AS AD_HRS,AD_SALARY,AD_SALCD, '" & txtWeek.Text & "' AS AD_WCBNBR FROM HR_ATTENDANCE"
'SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
'SQLQ = SQLQ & " GROUP BY AD_EMPNBR,AD_REASON,AD_CHRGCODE,AD_REGION,AD_SALARY,AD_SALCD"
'rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'If Not rsAttend.EOF Then
'End If
'rsAttend.Close
'Set rsAttend = Nothing

'xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xSFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
If glbSQL Then
    If ReportSel = "COS" Then
        SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    End If
End If

If Len(HisSQL) > 1 Then
    If ReportSel = "COS" Then
        If glbMulti Then
            SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
        Else
            SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
        End If
    End If
End If

SQLQ = SQLQ & " GROUP BY AD_EMPNBR,AD_REASON,AD_CHRGCODE,AD_REGION,AD_SALARY,AD_SALCD"
'Ticket #28539 - Only add Comments if Show Comments is checked otherwise the group by is showing an issue.
If chkComments Then
    SQLQ = SQLQ & " ,AD_COMM"      'Ticket #27892 - Added Comments
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
DoEvents

'If ReportSel = "HIS" Or chkInclHIS Then
    MDIMain.panHelp(0).FloodPercent = 60
    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & " SELECT " & Replace(xSFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
    If glbSQL Then
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
    End If
        
    If Len(HisSQL) > 1 Then
        If glbOracle Then
            SQLQ = SQLQ & " WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
        Else
            SQLQ = SQLQ & "WHERE (" & Replace(HisSQL, "AD_", "AH_") & ")"
        End If
        
    End If
    SQLQ = SQLQ & " GROUP BY AH_EMPNBR,AH_REASON,AH_CHRGCODE,AH_REGION,AH_SALARY,AH_SALCD"
    'Ticket #28539 - Only add Comments if Show Comments is checked otherwise the group by is showing an issue.
    If chkComments Then
        SQLQ = SQLQ & " ,AH_COMM"      'Ticket #27892 - Added Comments
    End If
    
    MDIMain.panHelp(0).FloodPercent = 65
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    DoEvents
    
'End If

'If chkCostHis = vbChecked And Not glbMulti Then
'    Dim rsIn As New ADODB.Recordset
'    Dim rsOUT As New ADODB.Recordset
'
'    If Not (glbOracle Or glbSQL) Then Pause (5)
'
'    SQLQ = "SELECT AD_ATT_ID, AD_SALARY, AD_SALCD, AD_WHRS, AD_EMPNBR, AD_DOA FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
'    rsOUT.Open SQLQ, gdbAdoIhr001W, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If rsOUT.EOF = False And rsOUT.BOF = False Then
'        Do '
'            'Historical Salary
'            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_TOTAL, SH_SALCD From HR_SALARY_HISTORY WHERE (SH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (SH_EDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY SH_EDATE DESC"
'            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'
'            If rsIn.EOF = False And rsIn.BOF = False Then
'                rsIn.MoveFirst
'                If glbCompSerial = "S/N - 2373W" Then
'                    rsOUT("AD_SALARY") = rsIn("SH_TOTAL")
'                Else
'                    rsOUT("AD_SALARY") = rsIn("SH_SALARY")
'                End If
'                rsOUT("AD_SALCD") = rsIn("SH_SALCD")
'            Else
'                rsOUT("AD_SALARY") = 0
'            End If
'            rsIn.Close
'
'            'Historical Position Weekly hours
'            SQLQ = "SELECT JH_SDATE, JH_WHRS From HR_JOB_HISTORY WHERE (JH_EMPNBR = " & rsOUT("AD_EMPNBR") & ") AND (JH_SDATE <= " & Date_SQL(rsOUT("AD_DOA")) & ") ORDER BY JH_SDATE DESC"
'            rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'            If rsIn.EOF = False And rsIn.BOF = False Then
'                rsIn.MoveFirst
'                rsOUT("AD_WHRS") = rsIn("JH_WHRS")
'            Else
'                rsOUT("AD_WHRS") = 0
'            End If
'            rsIn.Close
'            rsOUT.Update
'
'            rsOUT.MoveNext
'        Loop Until rsOUT.EOF
'    End If
'    rsOUT.Close
'    DoEvents
'    Set rsIn = Nothing
'    Set rsOUT = Nothing
'End If

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

