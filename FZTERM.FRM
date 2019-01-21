VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRTerm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Termination Report"
   ClientHeight    =   9615
   ClientLeft      =   255
   ClientTop       =   1455
   ClientWidth     =   14685
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
   ScaleHeight     =   9615
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   71
      Top             =   9120
      Width           =   12975
   End
   Begin VB.VScrollBar scrControl 
      Height          =   8775
      LargeChange     =   315
      Left            =   12960
      Max             =   100
      SmallChange     =   315
      TabIndex        =   70
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   12855
      Begin VB.CheckBox chkShowRetire 
         Caption         =   "Show Retirees too"
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
         Left            =   7080
         TabIndex        =   28
         Top             =   6900
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkShowSal 
         Caption         =   "Show Salary"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   6900
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   11
         Left            =   7200
         TabIndex        =   7
         Tag             =   "00-Cause for Termination"
         Top             =   2190
         Visible         =   0   'False
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TECA"
      End
      Begin VB.Frame frmEducation 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         TabIndex        =   72
         Top             =   5145
         Width           =   10695
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   9
            Left            =   2280
            TabIndex        =   22
            Tag             =   "00-Conducted by Code"
            Top             =   810
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ESCB"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   12
            Left            =   2280
            TabIndex        =   21
            Tag             =   "00-Course Type Code"
            Top             =   465
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCD"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   10
            Left            =   2280
            TabIndex        =   20
            Tag             =   "00-Course Type Code"
            Top             =   120
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCT"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   5
            Left            =   3840
            TabIndex        =   24
            Tag             =   "40-Completed Date upto and including this date"
            Top             =   1155
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   4
            Left            =   2280
            TabIndex        =   23
            Tag             =   "40-Completed Date from and including this date forward"
            Top             =   1155
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Type"
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
            TabIndex        =   76
            Top             =   165
            Width           =   900
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducted By"
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
            TabIndex        =   75
            Top             =   855
            Width           =   1005
         End
         Begin VB.Label lblRenew 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed Range"
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
            TabIndex        =   74
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Code"
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
            TabIndex        =   73
            Top             =   510
            Width           =   915
         End
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "First Level of grouping records"
         Top             =   7605
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Tag             =   "Second level of grouping records"
         Top             =   7920
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Tag             =   "Third level of grouping records"
         Top             =   8235
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Tag             =   "Final sorting of records - no totals"
         Top             =   8550
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2595
         MaxLength       =   4
         TabIndex        =   14
         Tag             =   "00-Employee Position Shift"
         Top             =   4170
         Visible         =   0   'False
         Width           =   450
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   6
         Tag             =   "00-Reason for Termination"
         Top             =   2190
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin Threed.SSOption optType 
         Height          =   195
         Index           =   1
         Left            =   2370
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "Detailed Report"
         Top             =   6900
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "   Detailed "
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
      Begin Threed.SSOption optType 
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   26
         Tag             =   "Summary Report"
         Top             =   6900
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "   Summary "
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
      Begin Threed.SSFrame fraReports 
         Height          =   1275
         Index           =   1
         Left            =   9840
         TabIndex        =   44
         Top             =   60
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   2249
         _StockProps     =   14
         Caption         =   "Summary Report Selection"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   2725
            _Version        =   65536
            _ExtentX        =   4807
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "     Health and Safety Total Cost"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Benefits Report"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Attendance Report"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Termination Report"
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
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1530
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
         Left            =   2280
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   1860
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
         Left            =   2280
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1200
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
         Left            =   2280
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   870
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   540
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
         Left            =   2280
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   210
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
         Left            =   2280
         TabIndex        =   12
         Tag             =   "00-Enter Administered By Code"
         Top             =   3510
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   13
         Tag             =   "00-Enter Section Code"
         Top             =   3840
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
         Left            =   2280
         TabIndex        =   11
         Tag             =   "00-Enter Region Code"
         Top             =   3180
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3990
         TabIndex        =   10
         Tag             =   "40-Date upto and including this date forward"
         Top             =   2850
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Tag             =   "40-Date from and including this date forward"
         Top             =   2850
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Tag             =   "10-Enter Employee Number"
         Top             =   2520
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
         Index           =   3
         Left            =   3960
         TabIndex        =   18
         Tag             =   "40-Date upto and including this date forward"
         Top             =   4860
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         Tag             =   "40-Date from and including this date forward"
         Top             =   4860
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Tag             =   "00-Enter Position Code"
         Top             =   4860
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   19
         Tag             =   "00-Attendance Reason Code"
         Top             =   5205
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   15
         Tag             =   "00-Enter Position Code"
         Top             =   4500
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDHC"
         MaxLength       =   6
      End
      Begin Threed.SSFrame fraReports 
         Height          =   1755
         Index           =   0
         Left            =   9840
         TabIndex        =   45
         Top             =   60
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   3096
         _StockProps     =   14
         Caption         =   "Detailed Report Selection"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Benefits"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Health and Safety"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "     Position/Salary/Performance"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Termination"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Attendance Report"
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
         Begin Threed.SSOption optReports 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "     Continuing Education"
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
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Term Cause"
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
         Left            =   6240
         TabIndex        =   77
         Top             =   2235
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   69
         Top             =   255
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
         TabIndex        =   68
         Top             =   585
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
         Left            =   120
         TabIndex        =   67
         Top             =   1245
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
         Left            =   120
         TabIndex        =   66
         Top             =   1575
         Width           =   450
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Reason"
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
         TabIndex        =   65
         Top             =   2235
         Width           =   1425
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
         Left            =   120
         TabIndex        =   64
         Top             =   2565
         Width           =   1290
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Term'n. From / To Date"
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
         Left            =   120
         TabIndex        =   63
         Top             =   2895
         Width           =   1665
      End
      Begin VB.Label lblSelCri 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   62
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblRepGrp 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   61
         Top             =   7370
         Width           =   1575
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
         Index           =   12
         Left            =   120
         TabIndex        =   60
         Top             =   6900
         Width           =   1065
      End
      Begin VB.Label lblTitle 
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
         Index           =   8
         Left            =   120
         TabIndex        =   59
         Top             =   7635
         Width           =   885
      End
      Begin VB.Label lblTitle 
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
         Index           =   9
         Left            =   120
         TabIndex        =   58
         Top             =   7950
         Width           =   885
      End
      Begin VB.Label lblTitle 
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
         Index           =   10
         Left            =   120
         TabIndex        =   57
         Top             =   8265
         Width           =   885
      End
      Begin VB.Label lblTitle 
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
         Index           =   11
         Left            =   120
         TabIndex        =   56
         Top             =   8580
         Width           =   660
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
         Left            =   120
         TabIndex        =   55
         Top             =   915
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
         TabIndex        =   54
         Top             =   3225
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
         TabIndex        =   53
         Top             =   3555
         Width           =   1125
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
         TabIndex        =   52
         Top             =   3885
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
         TabIndex        =   51
         Top             =   1905
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
         TabIndex        =   50
         Top             =   4215
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   49
         Top             =   4905
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance From / To Date"
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
         TabIndex        =   48
         Top             =   4905
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label lblAttReason 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Reason"
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
         TabIndex        =   47
         Top             =   5250
         Width           =   1425
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   4590
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   13320
      Top             =   8520
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
   Begin Crystal.CrystalReport vbxCrystal2 
      Left            =   13920
      Top             =   8520
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
   Begin Crystal.CrystalReport vbxCrystal3 
      Left            =   13320
      Top             =   9000
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
Attribute VB_Name = "frmRTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LocstrSelCri As String 'for wfc
Dim WFC_ACT_Flag As Boolean 'Ticket #15522

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Termination Criteria", Me) Then Exit Sub
    If Not Cri_SetAll() Then
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
'    cmdClose.SetFocus
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
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
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
'On Error GoTo CRW_Err

If CriCheck() Then
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    If Not Cri_SetAll() Then
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
'    cmdClose.SetFocus
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    If glbWFC Then
        If WFC_ACT_Flag Then
            Me.vbxCrystal2.WindowTitle = "Salary Continuance Employee Summary Report"
            Me.vbxCrystal2.Destination = 0
            Me.vbxCrystal2.Action = 1
            vbxCrystal2.Reset
        End If
        If chkShowRetire.Visible And chkShowRetire.Value Then 'Ticket #18772
            Me.vbxCrystal3.WindowTitle = "Employee Retirement Report"
            Me.vbxCrystal3.Destination = 0
            Me.vbxCrystal3.Action = 1
            vbxCrystal3.Reset
        End If
    End If
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
If Err = 20504 Then
Err = 0
Resume Next
End If

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
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "Term_HREMP.ED_LOC"
    If intIdx% = 1 Then strCd$ = "Term_HREMP.ED_ORG"
    If intIdx% = 2 Then strCd$ = "Term_HREMP.ED_EMP"
    If intIdx% = 3 Then strCd$ = "Term_HRTRMEMP.Term_Reason"
    If intIdx% = 4 Then strCd$ = "Term_HREMP.ED_REGION"
    If intIdx% = 5 Then strCd$ = "Term_HREMP.ED_ADMINBY"
    If intIdx% = 6 Then strCd$ = "Term_HREMP.ED_SECTION"
    If intIdx% = 7 Then strCd$ = "Term_ATTENDANCE.AD_REASON"
    If intIdx% = 8 Then strCd$ = "TERM_HREMP.ED_HIRECODE"
    If glbWFC Then
        If intIdx% = 11 Then strCd$ = "Term_HRTRMEMP.Term_Cause"
    End If
    
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "Term_HREMP.ED_REGION" Or strCd$ = "Term_HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
    
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    
    If glbWFC And Not (intIdx% = 3 Or intIdx% = 7 Or intIdx% = 11) Then
        'LocstrSelCri = glbstrSelCri
        LocstrSelCri = LocstrSelCri & " AND " & CodeCri 'Ticket #20072
    End If
    
    glbiOneWhere = True
End If


End Sub

Private Sub Cri_Dept()

Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String

If Len(clpDept.Text) > 0 Then
    DeptCri = "({Term_HREMP.ED_DEPTNO} in ['" & Replace(clpDept.Text, ",", "','") & "'])"
Else
    If glbNDepts < 0 Then    ' specific departments - less than has access to all of them
        ' they can see them all and not one selected
        Exit Sub
    End If
    countr = 0
    DeptCri = "({Term_HREMP.ED_DEPTNO} IN ['" & glbDepts(countr) & "'"
    'DeptCri = "({Term_HREMP.ED_DEPTNO} = '" & glbDepts(countr) & "'"
    countr = countr + 1
    Do While countr < glbNDepts
        DeptCri = DeptCri & ", '" & glbDepts(countr) & "'"
        countr = countr + 1
    Loop
    DeptCri = DeptCri & "])"
   ' DeptCri = DeptCri & ")"
End If

If Len(DeptCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DeptCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DeptCri
    End If
    glbiOneWhere = True
    If glbWFC Then
        'LocstrSelCri = glbstrSelCri
        LocstrSelCri = LocstrSelCri & " AND " & DeptCri 'Ticket #20072
    End If
End If


End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({Term_HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    If glbWFC Then
        'LocstrSelCri = glbstrSelCri
        LocstrSelCri = LocstrSelCri & " AND " & DivCri 'Ticket #20072
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{Term_HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    If glbWFC Then
        'LocstrSelCri = glbstrSelCri
        LocstrSelCri = LocstrSelCri & " AND " & EECri 'Ticket #20072
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({Term_HRTRMEMP.Term_DOT} "
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

'Ticket #24264
If Len(dlpDateRange(0).Text) > 0 Or Len(dlpDateRange(1).Text) > 0 Then
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({Term_HRTRMEMP.Term_DOT} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "({Term_HRTRMEMP.Term_DOT}  "
        TempCri = TempCri & " <= "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
End If

'Ticket #24264 - does not work
'For x% = 0 To 1
'    If Len(dlpDateRange(0).Text) > 0 Then
'        TempCri = "({Term_HRTRMEMP.Term_DOT}  "
'        If x% = 0 Then
'            TempCri = TempCri & " >= "
'        Else
'            TempCri = TempCri & " <= "
'        End If
'        dtYYY% = Year(dlpDateRange(0).Text)
'        dtMM% = month(dlpDateRange(0).Text)
'        dtDD% = Day(dlpDateRange(0).Text)
'        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'        GoTo Cri_FTDatst
'    End If
'Next x%

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

'Hemu - 07/22/2004 Begin - Ticket # 6549
Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String

If Len(clpJOB.Text) <= 0 Then Exit Sub

If (optType(0) And optReports(4)) Then
    PosCri = "({Term_SALARY_HISTORY.SH_JOB} = '" & clpJOB.Text & "')"
Else
    PosCri = "({Term_JOB_HISTORY.JH_JOB} = '" & clpJOB.Text & "')"
End If

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If

End Sub
'Hemu - 07/22/04 End

Private Sub Cri_FTAttDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If dlpDateRange(2).Visible = True And dlpDateRange(3).Visible = True Then
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        TempCri = "({Term_ATTENDANCE.AD_DOA} "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        
        dtYYY% = Year(dlpDateRange(3).Text)
        dtMM% = month(dlpDateRange(3).Text)
        dtDD% = Day(dlpDateRange(3).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTAttDatst
    End If
    
    'Ticket #24264
    If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
        If Len(dlpDateRange(2).Text) > 0 Then
            TempCri = "({Term_ATTENDANCE.AD_DOA} "
            TempCri = TempCri & " >= "
            dtYYY% = Year(dlpDateRange(2).Text)
            dtMM% = month(dlpDateRange(2).Text)
            dtDD% = Day(dlpDateRange(2).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTAttDatst
        End If
        If Len(dlpDateRange(3).Text) > 0 Then
            TempCri = TempCri & "({Term_ATTENDANCE.AD_DOA}  "
            TempCri = TempCri & " <= "
            dtYYY% = Year(dlpDateRange(3).Text)
            dtMM% = month(dlpDateRange(3).Text)
            dtDD% = Day(dlpDateRange(3).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTAttDatst
        End If
    End If
    
    'Ticket #24264 - does not work
'    For x% = 2 To 3
'        If Len(dlpDateRange(2).Text) > 0 Then
'            TempCri = "({Term_ATTENDANCE.AD_DOA}  "
'            If x% = 2 Then
'                TempCri = TempCri & " >= "
'            Else
'                TempCri = TempCri & " <= "
'            End If
'            dtYYY% = Year(dlpDateRange(2).Text)
'            dtMM% = month(dlpDateRange(2).Text)
'            dtDD% = Day(dlpDateRange(2).Text)
'            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'            GoTo Cri_FTAttDatst
'        End If
'    Next x%
End If

Cri_FTAttDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{Term_HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If

If glbWFC Then
    'LocstrSelCri = glbstrSelCri
    LocstrSelCri = LocstrSelCri & " AND " & EECri 'Ticket #20072
End If

glbiOneWhere = True

End Sub

Private Sub Cri_Conduct()
Dim EECri As String
Dim strCd$

If Len(clpCode(9).Text) < 1 Then Exit Sub

EECri = "({Term_HREDSEM.ES_CONDUCT} = '" & clpCode(9).Text & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_CourseType()
Dim EECri As String, OneSet%, x%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False

If Len(clpCode(10).Text) = 0 Then Exit Sub

strCa$ = "Term_HREDSEM.ES_CTYPE"

If Len(clpCode(10).Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} IN ['" & Replace(clpCode(10).Text, ",", "','") & "'])"
End If
    
If Len(EECri) > 0 Then
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

strCa$ = "Term_HREDSEM.ES_CRSCODE"

If Len(clpCode(12).Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} in ['" & Replace(clpCode(12).Text, ",", "','") & "'])"
End If

If Len(EECri) > 0 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_DateComp()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, fuTempCri As String

If Len(dlpDateRange(4).Text) > 0 And Len(dlpDateRange(5).Text) > 0 Then
    TempCri = "({Term_HREDSEM.ES_DATCOMP} "
    dtYYY% = Year(dlpDateRange(4).Text)
    dtMM% = month(dlpDateRange(4).Text)
    dtDD% = Day(dlpDateRange(4).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    
    dtYYY% = Year(dlpDateRange(5).Text)
    dtMM% = month(dlpDateRange(5).Text)
    dtDD% = Day(dlpDateRange(5).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(ES_DATCOMP >=" & Date_SQL(dlpDateRange(4).Text)
    fuTempCri = fuTempCri & " AND ES_DATCOMP <= " & Date_SQL(dlpDateRange(5).Text) & ")"
    GoTo Cri_FTDatst
End If

For x% = 4 To 5
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({Term_HREDSEM.ES_DATCOMP}  "
        If x% = 4 Then
            TempCri = TempCri & " >= "
            fuTempCri = "(ES_DATCOMP >= " & Date_SQL(dlpDateRange(4).Text) & ") "
        Else
            TempCri = TempCri & " <= "
            fuTempCri = "(ES_DATCOMP <= " & Date_SQL(dlpDateRange(5).Text) & ") "
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
    'DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim x%

Cri_SetAll = False

'On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
LocstrSelCri = "": WFC_ACT_Flag = False

Call glbCri_DeptUN(clpDept.Text)

glbstrSelCri = Replace(glbstrSelCri, "HREMP.", "Term_HREMP.")

If glbWFC Then LocstrSelCri = glbstrSelCri 'Ticket #15522
' call cri models set both glbiONeWhere and strSelCri

'Call Cri_Dept
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 8
    If x% = 7 Then
        If clpCode(7).Visible = True Then
            Call Cri_Code(x%)
        End If
    Else
        Call Cri_Code(x%)
    End If
Next x%
If glbWFC Then Call Cri_Code(11)

Call Cri_PT
Call Cri_Shift
Call Cri_EE

Call Cri_FTDates

'Hemu - 07/22/2004 Begin - Ticket # 6549
If clpJOB.Visible = True Then
    Call Cri_Position
End If
'Hemu - 07/22/2004 End

If clpCode(7).Visible = True Then
    Call Cri_FTAttDates
End If

If frmEducation.Visible = True Then
    Call Cri_CourseType
    Call Cri_CourseCode
    Call Cri_DateComp
    Call Cri_Conduct
End If

' optType(0)= Summary, (1)=Detail
If optType(0) = True Then
    If optReports(4).Value Then
        ' Termination Report
        Me.vbxCrystal.Formulas(9) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
        If glbSQL Then Me.vbxCrystal.Formulas(15) = "showsalary=" & IIf(chkShowSal, "True", "False") & ""
        Call setRptLabel(Me, 0)
        
        If glbWFC Then 'WFC asked Ticket #15522
            If chkShowSal Then
                If comGroup(0).Text <> "(none)" Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsum.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsu1.rpt"
                End If
            Else
                If comGroup(0).Text <> "(none)" Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsumws.rpt"
                    Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "rztrmsumwa.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsumw1.rpt"
                    Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "rztrmsumwa.rpt"
                End If
                WFC_ACT_Flag = True
            End If
            If chkShowRetire.Visible And chkShowRetire.Value Then   'Ticket #18772
                Me.vbxCrystal3.ReportFileName = glbIHRREPORTS & "RZtrmetire.rpt"
            End If
        Else
            If comGroup(0).Text <> "(none)" Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsum.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmsu1.rpt"
            End If
        End If
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
            If glbWFC Then
                If WFC_ACT_Flag Then
                    Me.vbxCrystal2.Connect = RptODBC_SQL
                End If
                If chkShowRetire.Visible And chkShowRetire.Value Then  'Ticket #18772
                    Me.vbxCrystal3.Connect = RptODBC_SQL
                End If
            End If
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            If comGroup(0).Text <> "(none)" Then
                For x% = 0 To 11
                    Select Case x%
                    Case 0 To 2
                        Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                    Case Else
                        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                    End Select
                Next x%
            Else
                For x% = 0 To 10
                    Select Case x%
                    Case 0 To 2, 10
                        Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                    Case Else
                        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                    End Select
                Next x%
            End If
        End If
        If glbNoNONE Then
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'NONE') "
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'NONE') "
                End If
            End If
        End If
        If glbNoEXEC Then    'Hemu -EXE
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'EXEC') "     'Hemu -EXE
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'EXEC') "  'Hemu -EXE
                End If
            End If
        End If
    End If
    If optReports(5).Value Then
        ' Attendance Report
        If comGroup(0).Text <> "(none)" Then
            If glbMulti Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmats_M.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmats.rpt"
            End If
        Else
            If glbMulti Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmat1_M.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmat1.rpt"
            End If
        End If
        If Len(dlpDateRange(2).Text) > 0 Then
            Me.vbxCrystal.Formulas(1) = "attenddate" & "= 'Attendance Date: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & "'"
        End If
        If Len(dlpDateRange(0).Text) > 0 Then
            Me.vbxCrystal.Formulas(2) = "termdate" & "= 'Termination Date: " & dlpDateRange(0).Text & " - " & dlpDateRange(1).Text & "'"
        End If
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 8 + 1
                Select Case x%
                Case 4 To 7
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
            Me.vbxCrystal.DataFiles(8) = glbIHRDB    'laura nov 26, 1997
        End If
    End If
    If optReports(6).Value Then
        ' Benefits Report
        If comGroup(0).Text <> "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmben.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmbe1.rpt"
        End If
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 8
                Select Case x%
                Case 4 To 6
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
        End If
    End If
    If optReports(7).Value Then
        MsgBox "Health and Safety Report is not ready yet"
        Exit Function
    End If
    
    ' set to sorting/grouping criteria
    x% = Cri_Sorts()   ' returns number of sections formated
Else
    If optReports(0).Value Then
        Call setRptLabel(Me, 1)
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmmst.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Formulas(3) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
            Me.vbxCrystal.Formulas(35) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
            Me.vbxCrystal.Formulas(36) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
            Me.vbxCrystal.Formulas(37) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "

            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 9 + 1
                Select Case x%
                Case 0 To 4
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
        End If
        If glbNoNONE Then
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'NONE') "
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'NONE') "
                End If
            End If
        End If
        If glbNoEXEC Then   'Hemu -EXE
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'EXEC') "     'Hemu -EXE
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'EXEC') "  'Hemu -EXE
                End If
            End If
        End If
        
    End If
    If optReports(1).Value Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmbnd.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 6 + 1
                Select Case x%
                Case 4 To 6
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
        End If
        x% = Cri_Sorts()   ' returns number of sections formated

    End If
    If optReports(2).Value Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmpos.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 8 + 1
                Select Case x%
                Case 0 To 3, 8
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
        End If
        If glbNoNONE Then
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'NONE') "
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'NONE') "
                End If
            End If
        End If
        If glbNoEXEC Then    'Hemu -EXE
            If Len(glbstrSelCri) > 0 Then
                If Not glbiOneWhere Then
                    glbstrSelCri = " ({Term_HREMP.ED_ORG } <> 'EXEC') "     'Hemu -EXE
                Else
                    glbstrSelCri = glbstrSelCri & " AND ({Term_HREMP.ED_ORG } <> 'EXEC') "  'Hemu -EXE
                End If
            End If
        End If
    End If
    If optReports(3).Value Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmini.rpt"
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 11 + 1
                Select Case x%
                Case 7 To 11
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
        End If
    End If
    
    'Hemu - 12/15/2003 Begin - Detailed Attendance Report
    If optReports(8).Value Then
        ' Attendance Report
        If comGroup(0).Text <> "(none)" Then
            If glbMulti Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmatd_M.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmatd.rpt"
            End If
        Else
            If glbMulti Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmad1_M.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmad1.rpt"
            End If
        End If
        If Len(dlpDateRange(2).Text) > 0 Then
            Me.vbxCrystal.Formulas(1) = "attenddate" & "= 'Attendance Date: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & "'"
        End If
        If Len(dlpDateRange(0).Text) > 0 Then
            Me.vbxCrystal.Formulas(2) = "termdate" & "= 'Termination Date: " & dlpDateRange(0).Text & " - " & dlpDateRange(1).Text & "'"
        End If
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            For x% = 0 To 8 + 1
                Select Case x%
                Case 4 To 7
                    Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
                Case Else
                    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
                End Select
            Next x%
            Me.vbxCrystal.DataFiles(8) = glbIHRDB    'laura nov 26, 1997
        End If
    End If
    'Hemu - 12/15/2003 End
    
    'Hemu - Ticket #13941 - Begin - Continuing Education Report
    If optReports(9).Value Then
        x% = Cri_Sorts()
        vbxCrystal.Formulas(40) = "lblCourseType='" & lStr("Course Type") & "'"
        vbxCrystal.Formulas(41) = "lblCourseCode='" & lStr("Course Code") & "'"
        vbxCrystal.Formulas(42) = "lblCompleted='" & lStr("Date Completed") & "'"
        vbxCrystal.Formulas(43) = "lblHours='" & lStr("Course Hours") & "'"

        If comGroup(0).Text <> "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmedu.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrmedu1.rpt"
        End If
        
        Me.vbxCrystal.Connect = RptODBC_SQL
    
        If Not glbSQL And Not glbOracle Then
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
            Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
            Me.vbxCrystal.DataFiles(3) = glbIHRDB
            Me.vbxCrystal.DataFiles(4) = glbIHRDB
            Me.vbxCrystal.DataFiles(5) = glbIHRAUDIT
            Me.vbxCrystal.DataFiles(6) = glbIHRDB
            Me.vbxCrystal.DataFiles(7) = glbIHRDB
            Me.vbxCrystal.DataFiles(8) = glbIHRDB
            Me.vbxCrystal.DataFiles(9) = glbIHRDB
            Me.vbxCrystal.DataFiles(10) = glbIHRDB
            Me.vbxCrystal.DataFiles(11) = glbIHRDB
            Me.vbxCrystal.DataFiles(12) = glbIHRDB
            Me.vbxCrystal.DataFiles(13) = glbIHRDB
            Me.vbxCrystal.DataFiles(14) = glbIHRAUDIT
            Me.vbxCrystal.DataFiles(15) = glbIHRAUDIT
        End If
        
    End If
    'Hemu - End
End If

If Len(glbstrSelCri) > 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    If glbWFC Then
        If WFC_ACT_Flag Then
            Me.vbxCrystal2.SelectionFormula = LocstrSelCri
        End If
        If chkShowRetire.Visible And chkShowRetire.Value Then  'Ticket #18772
            Me.vbxCrystal3.SelectionFormula = LocstrSelCri
        End If
    End If
Else
    Me.vbxCrystal.SelectionFormula = ""
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Terminated Employees Report"

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
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{Term_HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

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
Dim dscGroup$, GrpIdx%, Grp1, Grp2, Grp3

Cri_Sorts = 0
' first set primary grouping
If optType(0).Value = True And optReports(4).Value = True Then
    grpField$ = getEGroup(comGroup(0).Text)
    If grpField$ = "(none)" Then grpField$ = "{Term_HRTRMEMP.Company}"
    grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
    grpField$ = Replace(grpField$, "HREMP.", "Term_HREMP.")

    Grp1 = grpField$
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    If GrpIdx% <> 3 Then
        strSFormat$ = "GH1;T;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
        If grpField$ = "{Term_HRTRMEMP.Company}" Then
            strSFormat$ = "GF1;F;X;X;X;X;X;X"
        Else
            strSFormat$ = "GF1;T;X;X;X;X;X;X"
        End If
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
    Else
        strSFormat$ = "GH1;F;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
        strSFormat$ = "GF1;F;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
    End If
    
    If glbWFC And WFC_ACT_Flag Then 'Ticket #15522
        grpField$ = getEGroup(comGroup(0).Text)
        If grpField$ = "(none)" Then
            grpField$ = "{HREMP.ED_COMPNO}"
            Me.vbxCrystal2.Formulas(10) = "showgroup = false"
        End If
        grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
        grpField$ = Replace(grpField$, "HREMP.", "Term_HREMP.")
    
        Grp1 = grpField$
        dscGroup$ = comGroup(0).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal2.Formulas(0) = dscGroup$
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal2.GroupCondition(0) = grpCond$
        If GrpIdx% <> 3 Then
            strSFormat$ = "GH1;T;X;X;X;X;X;X"
            Me.vbxCrystal2.SectionFormat(z%) = strSFormat$
            z% = z% + 1
            strSFormat$ = "GF1;T;X;X;X;X;X;X"
            Me.vbxCrystal2.SectionFormat(z%) = strSFormat$
            z% = z% + 1
        Else
            strSFormat$ = "GH1;F;X;X;X;X;X;X"
            Me.vbxCrystal2.SectionFormat(z%) = strSFormat$
            z% = z% + 1
            strSFormat$ = "GF1;F;X;X;X;X;X;X"
            Me.vbxCrystal2.SectionFormat(z%) = strSFormat$
            z% = z% + 1
        End If
    End If
    If glbWFC And chkShowRetire.Visible And chkShowRetire.Value Then  'Ticket #18772
        grpField$ = getEGroup(comGroup(0).Text)
        If grpField$ = "(none)" Then
            grpField$ = "{HREMP.ED_COMPNO}"
            Me.vbxCrystal3.Formulas(10) = "showgroup = false"
        End If
        grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
        grpField$ = Replace(grpField$, "HREMP.", "Term_HREMP.")
    
        Grp1 = grpField$
        dscGroup$ = comGroup(0).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal3.Formulas(0) = dscGroup$
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal3.GroupCondition(0) = grpCond$
        If GrpIdx% <> 3 Then
            strSFormat$ = "GH1;T;X;X;X;X;X;X"
            Me.vbxCrystal3.SectionFormat(z%) = strSFormat$
            z% = z% + 1
            strSFormat$ = "GF1;T;X;X;X;X;X;X"
            Me.vbxCrystal3.SectionFormat(z%) = strSFormat$
            z% = z% + 1
        Else
            strSFormat$ = "GH1;F;X;X;X;X;X;X"
            Me.vbxCrystal3.SectionFormat(z%) = strSFormat$
            z% = z% + 1
            strSFormat$ = "GF1;F;X;X;X;X;X;X"
            Me.vbxCrystal3.SectionFormat(z%) = strSFormat$
            z% = z% + 1
        End If
    End If
    
    'Level 2
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HRTABL.TB_DESC}"
        Case 2: grpField$ = Grp1
    End Select
    Grp2 = grpField$
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    'Level 3
    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{HRTABL.TB_DESC}"
        Case 1: grpField$ = Grp2
    End Select
    Grp3 = grpField$
    grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(2) = grpCond$

    'Final Sort
    grpCond$ = "GROUP" & CStr(4) & ";" & "{Term_HRTRMEMP.Term_DOT}" & ";ANYCHANGE;D"
    Me.vbxCrystal.GroupCondition(3) = grpCond$

Else

    Y% = 0
    If Not glbLinamar Then ' Frank May 2, 2001
        Select Case comGroup(0).List(comGroup(0).ListIndex)
            Case lStr("Division"): grpField$ = "{HR_DIVISION.Division_Name}"
            Case lStr("Department"): grpField$ = "{HRDEPT.DF_NAME}"
            Case lStr("Location"): grpField$ = "{HRTABL1.TB_DESC}"
            Case "Employee Name": grpField$ = "{@EFullName}"
            Case lStr("Section"): grpField$ = "{tblSec.TB_DESC}"
            Case "Shift": grpField$ = "{Term_HREMP.ED_SHIFT}"
            Case lStr("Region"): grpField$ = "{tblRegion.TB_DESC}"
            Case lStr("Union"): grpField$ = "{tblUnion.TB_DESC}"
            Case "(none)": grpField$ = "(none)"
        End Select
    Else
        Select Case comGroup(0).List(comGroup(0).ListIndex)
            Case lStr("Division"): grpField$ = "{HR_DIVISION.Division_Name}"
            Case lStr("Department"): grpField$ = "{HRDEPT.DF_NAME}"
            Case lStr("Location"): grpField$ = "{HRTABL1.TB_DESC}"
            Case "Employee Name": grpField$ = "{@EFullName}"
            Case lStr("Section"): grpField$ = "{tblSec.TB_DESC}"
            Case "Employment Type": grpField$ = "{@EMPTYPE}"
            Case lStr("Region"): grpField$ = "{tblRegion.TB_DESC}"
            Case ("Home Line"): grpField$ = "{LN_HOMES.TB_DESC}"
            Case lStr("Union"): grpField$ = "{tblUnion.TB_DESC}"
            Case "(none)": grpField$ = "(none)"
        End Select
    End If
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
            Case 1: grpField$ = "{HRTABL.TB_DESC}"
            Case 2: grpField$ = "(none)"
        End Select
        ' items for x= 2 or less (groups 1,2,3 = x+ 1)

        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
            If comGroup(1).ListIndex <> 0 Then
                'Me.vbxCrystal.SectionFormat(0) = "GF4;F;X;X;X;X;X;X"
                'Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
            End If
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{HRTABL.TB_DESC}"
                Case 1: grpField$ = "(none)"
            End Select
            If grpField$ <> "(none)" Then
                grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(2) = grpCond$
            Else
                grpCond$ = "GROUP" & CStr(3) & ";" & "{Term_HRTRMEMP.Term_DOT}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(2) = grpCond$
            End If
        End If
    Else
        GrpIdx% = comGroup(1).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HRTABL.TB_DESC}"
            Case 2: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            dscGroup$ = ""
            'dscGroup$ = dscGroup1= ""
            dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
            Me.vbxCrystal.Formulas(0) = dscGroup$
            
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        Else
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{HRTABL.TB_DESC}"
                Case 1: grpField$ = "(none)"
            End Select
            If grpField$ <> "(none)" Then
                grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
            Else
                grpCond$ = "GROUP" & CStr(1) & ";" & "{Term_HRTRMEMP.Term_DOT}" & ";ANYCHANGE;D"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
            End If
        End If
    End If
End If

Cri_Sorts = z% ' next section number to format

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


For x% = 0 To 7
'If x% = 1 Then x% = x% + 1
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

'If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
'    MsgBox "If code entered it must be known"
'    clpCode(1).SetFocus
'    Exit Function
'End If


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

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

For x% = 2 To 3
    If Len(dlpDateRange(x%).Text) > 0 Then
        If Not IsDate(dlpDateRange(x%).Text) Then
            MsgBox "Not a valid date"
            dlpDateRange(x%).Text = ""
            dlpDateRange(x%).SetFocus
            Exit Function
        End If
    End If
Next x%

If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
    If DaysBetween(dlpDateRange(2), dlpDateRange(3)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(2).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

'Comment by Frank 04/21/2003
'If employee # entered, then system crashed
'Change the ocx file please Ticket# 4032
'If Not elpEEID.ListChecker Then
'    Exit Function
'End If
'
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

If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbWFC Then chkShowSal.Value = 0

Call INI_Controls(Me)

elpEEID.LookupType = TERM

Call optType_Click(1, 0)

lblTitle(1).Caption = lStr(lblTitle(1).Caption)

'If glbSQL Then
'    optReports(9).Visible = True
'Else
'    optReports(9).Visible = False
'End If

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJOB.TransDiv = glbWFCUserSecList
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
scrFrame.Height = 9255 '7815
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8500 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 6700 Then
            scrControl.Max = 5500
        Else
            scrControl.Max = 2000
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 10000 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7500 Then
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
Set frmRTerm = Nothing  'carmen apr 2000
End Sub

Private Sub optReports_Click(Index As Integer, Value As Integer)
    If (optReports(5) And optType(0)) Or (optReports(8) And optType(1)) Then
        lblTitle(0).Visible = True
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        lblAttReason.Visible = True
        clpCode(7).Visible = True
    Else
        lblTitle(0).Visible = False
        dlpDateRange(2).Visible = False
        dlpDateRange(3).Visible = False
        lblAttReason.Visible = False
        clpCode(7).Visible = False
        
    End If
    
    If glbWFC Then 'Ticket #15248
        If optType(0) Then
            If optReports(4) Then
                lblTitle(2).Visible = True
                clpCode(11).Visible = True
            Else
                lblTitle(2).Visible = False
                clpCode(11).Visible = False
                clpCode(11).Text = ""
            End If
        Else
            If optReports(0) Then
                lblTitle(2).Visible = True
                clpCode(11).Visible = True
            Else
                lblTitle(2).Visible = False
                clpCode(11).Visible = False
                clpCode(11).Text = ""
            End If
        End If
    End If

    'Ticket #15248
    If glbSQL Then
        If optType(0) And optReports(4) Then
            chkShowSal.Visible = True
        Else
            chkShowSal.Visible = False
        End If
    End If
    
    'Ticket #18772 for wfc pension system
    If glbWFC Then
        If optType(0) And optReports(4) Then
            chkShowRetire.Visible = True
        Else
            chkShowRetire.Visible = False
        End If
    End If
    
    'Hemu - Begin - Ticket # 6549
    If (optReports(4) And optType(0)) Or ((optReports(0) Or optReports(2)) And optType(1)) Then
        Label2.Visible = True
        clpJOB.Visible = True
    Else
        Label2.Visible = False
        clpJOB.Visible = False
    End If
    'Hemu - End
    
    'Hemu - Ticket #13941
    If optReports(9) And (fraReports(1).Visible = False) Then
        Label2.Visible = True
        clpJOB.Visible = True
        frmEducation.Visible = True
    Else
        frmEducation.Visible = False
    End If

End Sub

Private Sub optType_Click(Index As Integer, Value As Integer)
Dim I%

If optType(0) Then
    fraReports(1).Visible = True
    fraReports(0).Visible = False
    optReports(4).Value = True
    'panReportG.Visible = True 'js-01Apr99
    
    lblTitle(0).Visible = False
    dlpDateRange(2).Visible = False
    dlpDateRange(3).Visible = False
    clpCode(7).Visible = False
    lblAttReason.Visible = False
    If glbWFC Then
        If optReports(4) Then
            lblTitle(2).Visible = True
            clpCode(11).Visible = True
        Else
            lblTitle(2).Visible = False
            clpCode(11).Visible = False
        End If
    End If
    If glbSQL Then
        If optReports(4) Then
            chkShowSal.Visible = True
        Else
            chkShowSal.Visible = False
        End If
    End If
Else
    fraReports(1).Visible = False
    fraReports(0).Visible = True
    optReports(0).Value = True
    'panReportG.Visible = False 'js-01Apr99
    
    lblTitle(0).Visible = False
    dlpDateRange(2).Visible = False
    dlpDateRange(3).Visible = False
    clpCode(7).Visible = False
    lblAttReason.Visible = False
    If glbWFC Then
        If optReports(0) Then
            lblTitle(2).Visible = True
            clpCode(11).Visible = True
        Else
            lblTitle(2).Visible = False
            clpCode(11).Visible = False
        End If
    End If
    chkShowSal.Visible = False
End If

'Hemu - Begin - Ticket # 6549
If (optReports(4) And optType(0)) Or ((optReports(0) Or optReports(2)) And optType(1)) Then
    Label2.Visible = True
    clpJOB.Visible = True
Else
    Label2.Visible = False
    clpJOB.Visible = False
End If
'Hemu - End

'Hemu - Ticket #13941
If optReports(9) And (fraReports(1).Visible = False) Then
    Label2.Visible = True
    clpJOB.Visible = True
    frmEducation.Visible = True
Else
    frmEducation.Visible = False
End If

End Sub

Private Sub optType_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Region")
If glbWFC Then comGroup(0).AddItem lStr("Union") 'WFC Ticket #15522
comGroup(0).AddItem "(none)"
comGroup(1).AddItem "Employee Name"
comGroup(1).AddItem "Termination Reason"
comGroup(1).AddItem "(none)"
comGroup(2).AddItem "Termination Reason"
comGroup(2).AddItem "(none)"
comGroup(3).AddItem "Termination Date"
comGroup(0).ListIndex = 0
comGroup(1).ListIndex = 0
comGroup(2).ListIndex = 0
comGroup(3).ListIndex = 0
comGroup(3).Enabled = False
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
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
