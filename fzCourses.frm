VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRCourses 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Required Courses Report"
   ClientHeight    =   10950
   ClientLeft      =   435
   ClientTop       =   870
   ClientWidth     =   13215
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   13215
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   10695
      LargeChange     =   300
      Left            =   11760
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   55
      Top             =   120
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   10695
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   11415
      _Version        =   65536
      _ExtentX        =   20135
      _ExtentY        =   18865
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
         Height          =   8775
         Left            =   0
         ScaleHeight     =   8745
         ScaleWidth      =   11385
         TabIndex        =   32
         Top             =   120
         Width           =   11415
         Begin VB.OptionButton optEmpWork 
            Caption         =   "Employee Course History Report"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Tag             =   "Employee Work History Report"
            Top             =   6240
            Value           =   -1  'True
            Width           =   2715
         End
         Begin VB.OptionButton optCrossTrain 
            Caption         =   "Courses Required Report"
            Height          =   255
            Left            =   3375
            TabIndex        =   57
            Tag             =   "Cross-Training By Position Report"
            Top             =   6240
            Width           =   2325
         End
         Begin VB.CheckBox chkShowMedical 
            Caption         =   "Show Medical Contacts"
            Height          =   285
            Left            =   7320
            TabIndex        =   28
            Top             =   7320
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CheckBox chkForAudit 
            Caption         =   "For Data Audit"
            Height          =   285
            Left            =   7320
            TabIndex        =   29
            Top             =   7560
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkWeeklyEmpList 
            Caption         =   "Show Weekly Employee List"
            Height          =   285
            Left            =   7320
            TabIndex        =   27
            Top             =   7080
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Final sorting of records"
            Top             =   7575
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.CheckBox chkLastDay 
            Caption         =   "Show Last Day"
            Height          =   285
            Left            =   7320
            TabIndex        =   26
            Top             =   6840
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtShift 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2050
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "00-Employee Position Shift"
            Top             =   4530
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Final sorting of records"
            Top             =   7545
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "First Level of grouping records"
            Top             =   7200
            Visible         =   0   'False
            Width           =   2325
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   1740
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
            Left            =   1740
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
            Left            =   1740
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
            Left            =   1740
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
            Left            =   1740
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
            TabIndex        =   12
            Tag             =   "00-Enter Administered By Code"
            Top             =   3540
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
            Left            =   1740
            TabIndex        =   13
            Tag             =   "00-Enter Section Code"
            Top             =   3870
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
            TabIndex        =   11
            Tag             =   "00-Enter Region Code"
            Top             =   3210
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
            TabIndex        =   10
            Tag             =   "40-Position Start Date upto and including this date forward"
            Top             =   2880
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
            TabIndex        =   9
            Tag             =   "40-Position Start Date from and including this date forward"
            Top             =   2880
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   1740
            TabIndex        =   6
            Tag             =   "10-Enter Employee Number"
            Top             =   2220
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
            Left            =   8970
            TabIndex        =   24
            Tag             =   "40-Date upto and including this date forward"
            Top             =   7905
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   2
            Left            =   7140
            TabIndex        =   23
            Tag             =   "40-Date from and including this date forward"
            Top             =   7905
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpGrid 
            Height          =   285
            Left            =   8040
            TabIndex        =   8
            Top             =   2550
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
            Left            =   6600
            TabIndex        =   25
            Tag             =   "00-Benefit - Group Code"
            Top             =   6540
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "BGMF"
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   14
            Tag             =   "10-Reporting Authority 1"
            Top             =   4200
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   1
            Left            =   3660
            TabIndex        =   15
            Tag             =   "10-Reporting Authority 2"
            Top             =   4200
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   2
            Left            =   5580
            TabIndex        =   16
            Tag             =   "10-Reporting Authority 3"
            Top             =   4200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.DateLookup dlpAsOf 
            Height          =   285
            Left            =   1740
            TabIndex        =   20
            Tag             =   "40-As of Date Renewal Date"
            Top             =   5530
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpCrsCode 
            Height          =   285
            Left            =   1740
            TabIndex        =   18
            Tag             =   "00-Course Code"
            Top             =   4875
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowDescription =   0   'False
            TABLName        =   "ESCD"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   3
            Left            =   1740
            TabIndex        =   19
            Tag             =   "00-Course Type Code"
            Top             =   5200
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "ESCT"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            Height          =   285
            Left            =   1740
            TabIndex        =   7
            Tag             =   "00-Enter Position Code "
            Top             =   2550
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
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
            TabIndex        =   61
            Top             =   5245
            Width           =   1260
         End
         Begin VB.Label lblCrsCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   4915
            Width           =   915
         End
         Begin VB.Label lblBenGroup 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   56
            Top             =   6585
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblGrp 
            BackStyle       =   0  'Transparent
            Caption         =   "Work History Sort"
            Height          =   375
            Index           =   1
            Left            =   5520
            TabIndex        =   54
            Top             =   7605
            Visible         =   0   'False
            Width           =   1695
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
            TabIndex        =   53
            Top             =   2580
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblEmplStFrpmTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status From / To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   52
            Top             =   7920
            Visible         =   0   'False
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
            TabIndex        =   51
            Top             =   6960
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label lblShift 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   4575
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
            TabIndex        =   49
            Top             =   1935
            Width           =   630
         End
         Begin VB.Label lblRep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   4245
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
            TabIndex        =   47
            Top             =   3915
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
            TabIndex        =   46
            Top             =   3585
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
            TabIndex        =   45
            Top             =   3255
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
            TabIndex        =   44
            Top             =   945
            Width           =   615
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
            TabIndex        =   43
            Top             =   7545
            Visible         =   0   'False
            Width           =   660
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
            TabIndex        =   42
            Top             =   7230
            Visible         =   0   'False
            Width           =   885
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
            TabIndex        =   41
            Top             =   6960
            Visible         =   0   'False
            Width           =   1575
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
            TabIndex        =   40
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblFromTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Start Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   2925
            Width           =   1320
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
            TabIndex        =   38
            Top             =   2595
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
            TabIndex        =   37
            Top             =   2265
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
            TabIndex        =   36
            Top             =   1605
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
            TabIndex        =   35
            Top             =   1275
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
            TabIndex        =   34
            Top             =   615
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
            TabIndex        =   33
            Top             =   285
            Width           =   555
         End
         Begin VB.Label lblAsOf 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "As of Renewal Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   5575
            Width           =   1425
         End
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   12360
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
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, SQLQ
Dim rsTabl As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
    Dim x%
    
    On Error GoTo PrntErr
    
    If CriCheck() Then
        If Not PrtForm("Required Courses Report Criteria", Me) Then Exit Sub
        
        Call set_PrintState(False)
        
        x% = Cri_SetAll()
        
        Me.vbxCrystal.Destination = 1
        MDIMain.Timer1.Enabled = False
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        MDIMain.Timer1.Enabled = True
        
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
    Dim x%
    Dim strWHand As String
    On Error GoTo CRW_Err
    
    If CriCheck() Then
        Screen.MousePointer = HOURGLASS
        Call set_PrintState(False)
        
        'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
        'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
        Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
        
        x% = Cri_SetAll()
        
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
    Resume Next
    Screen.MousePointer = DEFAULT
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")  'Jaddy jun 16,1999
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Position Code"
    comGroup(0).AddItem lStr("Machine #")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0
    
    comGroup(2).AddItem "Descending"
    comGroup(2).AddItem "Ascending"
    comGroup(2).ListIndex = 0
End Sub

Private Sub Cri_Assoc()
    Dim EECri As String
    
    If Len(clpCode(1).Text) <= 0 Then Exit Sub
    
    If glbMulti And optCrossTrain Then
        EECri = "{qry_Employee_Work_History.JH_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    Else
        EECri = "{HREMP.ED_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_Dept()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim DeptCri As String
    
    DeptCri = ""
    
    Call glbCri_DeptUN(clpDept.Text)
End Sub

Private Sub Cri_Div()
    Dim DivCri As String
    
    If Len(clpDiv.Text) <= 0 Then Exit Sub
    
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    Else
        glbstrSelCri = DivCri
    End If
End Sub

Private Sub Cri_EE()
    Dim EECri As String
    
    If Len(elpEEID.Text) <= 0 Then Exit Sub
    
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_RepAuth()
    Dim TempCri As String
    Dim EECri As String, LocCri As String
    Dim I, xTemp As Boolean
    
    xTemp = False
    EECri = ""
    
    If optEmpWork Then Exit Sub

    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & "{qry_Employee_Work_History.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND {qry_Employee_Work_History.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        Else
            EECri = EECri & "{qry_Employee_Work_History.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND {qry_Employee_Work_History.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        Else
            EECri = EECri & "{qry_Employee_Work_History.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
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
    Dim x%
    Dim EECri As String, LocCri As String
    
    If optEmpWork Then Exit Sub
    If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "({qry_Employee_Work_History.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({qry_Employee_Work_History.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "({qry_Employee_Work_History.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
        GoTo Cri_FTDatst
    End If
    
    For x% = 0 To 1
        If Len(dlpDateRange(0).Text) > 0 Then
            TempCri = "({qry_Employee_Work_History.JH_SDATE}  "
            If x% = 0 Then
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
    Next x%

Cri_FTDatst:
    If Len(TempCri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & TempCri
        Else
          glbstrSelCri = TempCri
        End If
    End If
End Sub

Private Sub Cri_Position()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim PosCri As String
    
    If Len(clpJOB.Text) <= 0 Then Exit Sub
    If optEmpWork Then Exit Sub
        
    PosCri = "({qry_Employee_Work_History.JH_JOB} = '" & clpJOB.Text & "')"
    
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
    If optEmpWork Then Exit Sub
    
    GirdCri = "({qry_Employee_Work_History.JH_GRID} = '" & clpGrid.Text & "')"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & GirdCri
    Else
        glbstrSelCri = GirdCri
    End If
End Sub

Private Sub Cri_PT()
    Dim EECri As String
    
    If Len(clpPT.Text) < 1 Then Exit Sub
    
    If glbMulti And optCrossTrain Then
        EECri = "{qry_Employee_Work_History.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
    Else
        EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_CourseType()
    Dim EECri As String

    If Len(clpCode(3).Text) < 1 Then Exit Sub
    
    EECri = "{tblCourse.TB_USR1} in ['" & Replace(clpCode(3).Text, ",", "','") & "']"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_CourseCode()
    Dim EECri As String

    If Len(clpCrsCode.Text) < 1 Then Exit Sub
    
    EECri = "{HRCRSHIST_WRK.TR_CRSCODE} in ['" & Replace(clpCrsCode.Text, ",", "','") & "']"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
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

Private Function Cri_SetAll()
    Dim x%, strRName$, I
    
    On Error GoTo modSetCriteria_Err
    
    Cri_SetAll = False
    
    Screen.MousePointer = HOURGLASS
    
    glbiOneWhere = False
    glbstrSelCri = ""
    SQLQ = ""

    ' call cri models set both glbiONeWhere and strSelCri
    Call glbCri_DeptUN(clpDept.Text)
    SQLQ = glbstrSelCri
    Call Cri_Div
    Call Cri_Assoc
    Call Cri_Code(0)
    Call Cri_Code(1)
    Call Cri_Code(2)
    Call Cri_PT
    Call Cri_EE
    Call Cri_Position
    'Call Cri_Grid
    Call Cri_FTDates
    Call Cri_Status
    Call Cri_Code(7)
    Call Cri_Code(8)
    Call Cri_Code(9)
    Call Cri_RepAuth
    Call Cri_Shift
    
    Call Cri_CourseType
    If optEmpWork And Len(clpCrsCode.Text) > 0 Then
        Call Cri_CourseCode
    End If
        
    If optCrossTrain Then
        If Len(clpCrsCode.Text) > 0 Then
            Me.vbxCrystal.GroupSelectionFormula = "({HR_TRAIN.TR_CRSCODE} in ['" & Replace(clpCrsCode.Text, ",", "','") & "'])"
        End If
    End If
    
    'X% = Cri_Sorts()
    
    If optEmpWork Then
        'Call procedure to populate the temporary table for the report
        Call Populate_Course_History_Work_Table(SQLQ)
        
        glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HRCRSHIST_WRK.TR_WRKEMP}='" & glbUserID & "'"
    End If
    
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    
    If IsDate(dlpAsOf.Text) And optCrossTrain Then
        Me.vbxCrystal.Formulas(1) = "AsOfDate=Date(" & Year(dlpAsOf.Text) & ", " & month(dlpAsOf.Text) & ", " & Day(dlpAsOf.Text) & ")"
    End If
    
    If optEmpWork Then
        If glbCompSerial = "S/N - 2279W" Then  'Friesens
            strRName$ = glbIHRREPORTS & "SN2279_rzCourseHist.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzCourseHist.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Employee Course History Report"
    End If
    If optCrossTrain Then
        If glbCompSerial = "S/N - 2279W" Then  'Friesens
            strRName$ = glbIHRREPORTS & "SN2279_rzCourseReq.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzCourseReq.rpt"
        End If
        Me.vbxCrystal.WindowTitle = "Courses Required Report"
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
    
    Cri_SetAll = True

    Screen.MousePointer = DEFAULT

Exit Function
modSetCriteria_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select Report Criteria", "Courses Report", "Select")
    Cri_SetAll = False
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Function

Private Function Cri_Sorts()
    Dim grpCond$, grpField$
    Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
    Dim dscGroup$, GrpIdx%
    
    Cri_Sorts = 0
    
    'first set primary grouping
    z% = 0
    x% = 0
    
    grpField$ = getEGroup(comGroup(0).Text)
    Y% = x% + 1
    
    If comGroup(0) = "(none)" Then grpField$ = "{@EFullName}"
    Call setRptLabel(Me, 0)
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(x%) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(x%) = grpCond$
    
    Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
    Dim EECri As String, LocCri As String
    
    If Len(clpCode(2).Text) <= 0 Then Exit Sub
    
    If Len(clpCode(2).Text) > 0 Then
        EECri = "{HREMP.ED_EMP} = '" & clpCode(2).Text & "' "
    End If
    
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
        glbiOneWhere = True
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
    
        If Len(strCd$) > 0 Then
            CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
            If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
                CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
            End If
            If Len(glbstrSelCri) > 1 Then
                glbstrSelCri = glbstrSelCri & " AND " & CodeCri
            Else
                glbstrSelCri = CodeCri
            End If
        End If
    End If
End Sub

Private Function CriCheck()
    Dim x%, I
    
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
        
    For x% = 0 To 2
        If Not clpCode(x).ListChecker Then Exit Function
    Next x%
    For x% = 7 To 9
        If Not clpCode(x).ListChecker Then Exit Function
    Next x%
    
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
            MsgBox "'To' Position Start Date cannot be prior to 'From' Date"                       '
            Me.dlpDateRange(0).SetFocus                                         '
            Exit Function                                                       '
        End If
    End If
    
    For I = 0 To 2
        If elpRept(I).Caption = "Enter Valid Employee #" Then
            MsgBox "If Reporting Authority Entered - they must exist"
            elpRept(I).SetFocus
            Exit Function
        End If
    Next
    
    If Not elpEEID.ListChecker Then
        Exit Function
    End If
    
    If Not clpCrsCode.CheckList Then Exit Function
    
    If Not clpCode(3).CheckList Then Exit Function
    
    CriCheck = True
    
End Function

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    glbOnTop = Me.name
    
    If glbMultiGrid Then
        lblGrid.Visible = True
        clpGrid.Visible = True
    End If
    
    If Not glbMulti Then
        lblShift.Visible = True
        txtShift.Visible = True
    End If
    
    Call setRptCaption(Me)
    Call comGrpLoad
    
'    If Me.Caption = "Employee Profile Report" Then
'        lblGrp(1).Visible = True
'        comGroup(2).Visible = True
'    Else
'        lblGrp(1).Visible = False
'        comGroup(2).Visible = False
'    End If
    
    If glbLinamar Then clpCode(7).MaxLength = 8
    If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
    If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
    
    dlpAsOf.Text = Date
    
    If optEmpWork Then
        dlpAsOf.Visible = False
        lblAsOf.Visible = False
    Else
        dlpAsOf.Visible = True
        lblAsOf.Visible = True
    End If
    
    Call INI_Controls(Me)

    If glbWFC Then 'Ticket #25911 Franks 10/21/2014
        clpJOB.TransDiv = glbWFCUserSecList
    End If

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
On Error GoTo EH
    Dim c As Long
    
    If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
        panWindow.Height = Me.ScaleHeight - 200
        panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
        If panWindow.Height >= 7500 Then   '+ 230 Then
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
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form Resize", "Courses Required Rpt", "Form Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmRCourses = Nothing  'carmen apr 2000
End Sub

Private Sub optCrossTrain_Click()
    If optCrossTrain Then
        lblCrsCode.Visible = True
        clpCrsCode.Visible = True
        dlpAsOf.Visible = True
        lblAsOf.Visible = True
    Else
        'lblCrsCode.Visible = False
        'clpCrsCode.Visible = False
        dlpAsOf.Visible = False
        lblAsOf.Visible = False
    End If
End Sub

Private Sub optEmpWork_Click()
    If optEmpWork Then
        'lblCrsCode.Visible = False
        'clpCrsCode.Visible = False
        dlpAsOf.Visible = False
        lblAsOf.Visible = False
    Else
        dlpAsOf.Visible = True
        lblAsOf.Visible = True
    End If
    
End Sub

Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
End Sub

Private Sub txtShift_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Cri_Shift()
    Dim EECri As String, OneSet%, x%
    
    If Len(txtShift.Text) < 1 Then Exit Sub
        
    EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
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
    Dim dtYYY%, dtMM%, dtDD%, x%
    Dim FromDate, ToDate, SQLQ
    Dim RsHRPARCO As New ADODB.Recordset
    
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

Private Function GetJobCodeDesc(xKey)
    Dim rsTabl As New ADODB.Recordset
    Dim SQLQ As String, xStr As String
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xKey & "' "
    rsTabl.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTabl.EOF Then
        xStr = rsTabl("JB_DESCR")
    End If
    rsTabl.Close
    
    GetJobCodeDesc = Left(xStr, 20)
End Function

Private Function GetTABLDesc(xName, xKey)
    Dim rsTabl As New ADODB.Recordset
    Dim SQLQ As String, xStr As String
    
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xName & "' AND TB_KEY = '" & xKey & "' "
    rsTabl.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTabl.EOF Then
        xStr = rsTabl("TB_DESC")
    End If
    rsTabl.Close
    GetTABLDesc = Left(xStr, 20)
End Function

Private Function Populate_Course_History_Work_Table(xDeptSQL)
    Dim rsCrsHis As New ADODB.Recordset
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim xLstEmpNo As Long
    Dim SQLQ, xWhere, xDept, xJWhere, xDWhere, xLstCourse, xPosType As String
    Dim sDate

    
    'Delete the existing records of this user in this temp. table
    gdbAdoIhr001.BeginTrans
    SQLQ = "DELETE FROM HRCRSHIST_WRK " & in_SQL(glbIHRDBW) & " WHERE TR_WRKEMP = '" & glbUserID & "'"
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    
    SQLQ = "SELECT * FROM HRCRSHIST_WRK WHERE 1 = 2"
    rsCrsHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsCrsHis.EOF Then
        'Get the records from the Training List
        'Courses Not Taken ever or Taken & Renewed
        xWhere = ""
        SQLQ = "SELECT * FROM HR_TRAIN WHERE TR_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE "
        
        'Division
        If Len(clpDiv.Text) > 0 Then xWhere = "HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
        
        'Department
        'Call glbCri_DeptUN(clpDept.Text)
        'xDept = Replace(Replace(glbstrSelCri, "[", "("), "]", ")")
        'If Len(xDept) > 0 Then
        '    If Len(xWhere) > 0 Then
        '        xWhere = xWhere & " AND "
        '    End If
        '    xWhere = xWhere & xDept
        'End If
        xDeptSQL = Replace(Replace(xDeptSQL, "[", "("), "]", ")")
        If Len(xDeptSQL) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & xDeptSQL
        End If
        
        'Location
        If Len(clpCode(0).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
        End If
            
        'Union
        If Len(clpCode(1).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
        End If

        'Status
        If Len(clpCode(2).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
        End If
        
        'Category
        If Len(clpPT.Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"
        End If
        
        'Employee #
        If Len(elpEEID.Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
        End If
        
        'Region
        If Len(clpCode(7).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_REGION IN ('" & Replace(clpCode(7).Text, ",", "','") & "')"
        End If
        
        'Administered By
        If Len(clpCode(8).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_ADMINBY IN ('" & Replace(clpCode(8).Text, ",", "','") & "')"
        End If
        
        'Section
        If Len(clpCode(9).Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_SECTION IN ('" & Replace(clpCode(9).Text, ",", "','") & "')"
        End If
        
        'Shift
        If Len(txtShift.Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HREMP.ED_SHIFT = '" & txtShift.Text & "'"
        End If
        
        xWhere = xWhere & ")"
        
        'Position
        If Len(clpJOB.Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HR_TRAIN.TR_JOB = '" & clpJOB.Text & "'"
        End If
        
        'From/To Date
        xDWhere = ""
        If Len(dlpDateRange(0).Text) > 0 Or Len(dlpDateRange(1).Text) > 0 Then
            'If Len(xWhere) > 0 Then
            '    xWhere = xWhere & " AND "
            'End If
            
            If IsDate(dlpDateRange(0).Text) Then
                xDWhere = "HR_TRAIN.TR_SDATE >= " & Date_SQL(dlpDateRange(0).Text)
                'xWhere = xWhere & "HR_TRAIN.TR_SDATE >= '" & Date_SQL(dlpDateRange(0).Text) & "'"
            End If
            If IsDate(dlpDateRange(1).Text) Then
                If Len(xDWhere) > 0 Then
                    xDWhere = xDWhere & " AND "
                End If
                xDWhere = xDWhere & "HR_TRAIN.TR_SDATE <= " & Date_SQL(dlpDateRange(1).Text)
                'xWhere = xWhere & "HR_TRAIN.TR_SDATE <= '" & Date_SQL(dlpDateRange(1).Text) & "'"
            End If
        End If
                
        'Course Code
        If Len(clpCrsCode.Text) > 0 Then
            If Len(xWhere) > 0 Then
                xWhere = xWhere & " AND "
            End If
            xWhere = xWhere & "HR_TRAIN.TR_CRSCODE = '" & clpCrsCode.Text & "'"
        End If
                
        'Reporthing Authority 1,2,3
        xJWhere = ""
        If Len(Trim(elpRept(0).Text)) > 0 Or Len(Trim(elpRept(1).Text)) > 0 Or Len(Trim(elpRept(2).Text)) > 0 Then
            'If Len(xDWhere) > 0 Then
            '    xDWhere = xDWhere & " AND "
            'ElseIf Len(xWhere) > 0 Then
            '    xWhere = xWhere & " AND "
            'End If
            
            If Len(Trim(elpRept(0).Text)) > 0 Then
                xJWhere = " TR_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE "
                xJWhere = xJWhere & "HR_JOB_HISTORY.JH_REPTAU = " & Trim(elpRept(0).Text) & " "
            End If
            If Len(Trim(elpRept(1).Text)) > 0 Then
                If Len(xJWhere) > 0 Then
                    xJWhere = xJWhere & " AND "
                Else
                    xJWhere = xJWhere & " TR_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE "
                End If
                xJWhere = xJWhere & "HR_JOB_HISTORY.JH_REPTAU2 = " & Trim(elpRept(1).Text) & " "
            End If
            If Len(Trim(elpRept(2).Text)) > 0 Then
                If Len(xJWhere) > 0 Then
                    xJWhere = xJWhere & " AND "
                Else
                    xJWhere = xJWhere & " TR_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE "
                End If
                xJWhere = xJWhere & "HR_JOB_HISTORY.JH_REPTAU3 = " & Trim(elpRept(2).Text) & " "
            End If
            'xJWhere = xJWhere & ")"
            
            'xWhere = xWhere & xJWhere
        End If
        
        SQLQ = SQLQ & xWhere & IIf(Len(xDWhere) > 0, " AND " & xDWhere, xDWhere) & IIf(Len(xJWhere) > 0, " AND " & xJWhere, xJWhere)
        
        If Len(xJWhere) > 0 Then SQLQ = SQLQ & ")"
        
        SQLQ = SQLQ & " ORDER BY TR_EMPNBR, TR_CRSCODE, TR_JOB, TR_SDATE DESC"
        SQLQ = Replace(Replace(SQLQ, "{", "("), "}", ")")
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            rsHRTrain.MoveFirst
            Do While Not rsHRTrain.EOF
                rsCrsHis.AddNew
                rsCrsHis("TR_COMPNO") = "001"
                rsCrsHis("TR_EMPNBR") = rsHRTrain("TR_EMPNBR")
                rsCrsHis("TR_CRSCODE") = rsHRTrain("TR_CRSCODE")
                rsCrsHis("TR_JOB") = rsHRTrain("TR_JOB")
                rsCrsHis("TR_SDATE") = rsHRTrain("TR_SDATE")
                rsCrsHis("TR_POS_TYPE") = rsHRTrain("TR_POS_TYPE")
                rsCrsHis("TR_RENEW") = rsHRTrain("TR_RENEW")
                rsCrsHis("TR_COURSE_TAKEN") = rsHRTrain("TR_COURSE_TAKEN")
                rsCrsHis("TR_WRKEMP") = glbUserID
                'rsCrsHis ("TR_ENDDATE")
                rsCrsHis.Update
                
                rsHRTrain.MoveNext
            Loop
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
        
        'Get the records from Continuing Education table
        'Course Taken and Not Renewed
        xLstCourse = ""
        xLstEmpNo = 0
        
        If Len(xDWhere) > 0 Then
            xDWhere = Replace(Replace(xDWhere, "HR_TRAIN", "HR_JOB_HISTORY"), "TR_", "JH_")
        End If
        If Len(xJWhere) > 0 Then
            xJWhere = Replace(xJWhere, "TR_", "ES_")
        ElseIf Len(xDWhere) > 0 Then
            xDWhere = " ES_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE " & xDWhere
        End If
        xWhere = Replace(Replace(xWhere, "HR_TRAIN", "HREDSEM"), "TR_", "ES_")
        
        SQLQ = "SELECT * FROM HREDSEM WHERE (ES_RENEW IS NULL OR ES_RENEW  = '') "
        SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE "
        
        SQLQ = SQLQ & xWhere & IIf(Len(xJWhere) > 0, " AND " & xJWhere, xJWhere) & IIf(Len(xDWhere) > 0, " AND " & xDWhere, xDWhere)
        If Len(xJWhere) > 0 Or Len(xDWhere) > 0 Then SQLQ = SQLQ & ")"
        
        SQLQ = SQLQ & " ORDER BY ES_EMPNBR, ES_CRSCODE, ES_JOB, ES_DATCOMP DESC"
        SQLQ = Replace(Replace(SQLQ, "{", "("), "}", ")")
        
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            rsContEdu.MoveFirst
            Do While Not rsContEdu.EOF
                If xLstEmpNo = rsContEdu("ES_EMPNBR") And xLstCourse = rsContEdu("ES_CRSCODE") Then
                    'Same course taken multiple times, skip to next record
                    GoTo Next_ContEdu_Rec
                End If
                
                xLstEmpNo = rsContEdu("ES_EMPNBR")
                xLstCourse = rsContEdu("ES_CRSCODE")
                
                rsCrsHis.AddNew
                rsCrsHis("TR_COMPNO") = "001"
                rsCrsHis("TR_EMPNBR") = rsContEdu("ES_EMPNBR")
                rsCrsHis("TR_CRSCODE") = rsContEdu("ES_CRSCODE")
                rsCrsHis("TR_JOB") = rsContEdu("ES_JOB")
                
                sDate = get_Job_History_Data(rsContEdu("ES_EMPNBR"), rsContEdu("ES_JOB"), "JH_SDATE")
                If IsDate(sDate) Then
                    rsCrsHis("TR_SDATE") = CVDate(sDate)
                End If
                
                xPosType = get_Job_History_Data(rsContEdu("ES_EMPNBR"), rsContEdu("ES_JOB"), "POS_TYPE")
                rsCrsHis("TR_POS_TYPE") = xPosType
                
                'rsCrsHis("TR_RENEW") = rsContEdu("TR_RENEW")
                rsCrsHis("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")
                rsCrsHis("TR_WRKEMP") = glbUserID
                'rsCrsHis ("TR_ENDDATE")
                
                rsCrsHis.Update
Next_ContEdu_Rec:
                rsContEdu.MoveNext
            Loop
        End If
        rsContEdu.Close
        Set rsContEdu = Nothing
    End If
    
    rsCrsHis.Close
    Set rsCrsHis = Nothing
    
End Function

Private Function get_Job_History_Data(xEmpNo, xJob, xField)
    Dim rsJobHis As New ADODB.Recordset
    Dim SQLQ As String
    Dim xAsField
    
    'Check in both job tables which position record best matches the criteria and retrieve the most recent record.
    'Primary Current record is at the top, followed by Temp. Current and then tracked Previous, un-Tracked Previous
    xAsField = Replace(xField, "JH_", "TW_")
    If InStr(xAsField, "TW_") <> 0 Then
        SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, " & xField & " AS " & xAsField & ", 'C' AS POS_TYPE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL "
    Else
        SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'C' AS POS_TYPE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL "
    End If
    SQLQ = SQLQ & " FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " UNION "
    If InStr(xAsField, "TW_") <> 0 Then
        SQLQ = SQLQ & "SELECT TW_EMPNBR, " & xAsField & ", 'T' AS POS_TYPE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL "
    Else
        SQLQ = SQLQ & "SELECT TW_EMPNBR, 'T' AS POS_TYPE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL "
    End If
    SQLQ = SQLQ & " FROM HR_TEMP_WORK "
    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
    rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJobHis.EOF Then
        'Get the data of the top record on the list (most recent)
        rsJobHis.MoveFirst
        If InStr(xAsField, "TW_") <> 0 Then
            get_Job_History_Data = rsJobHis(xAsField)
        Else
            If Not IsNull(rsJobHis("TW_TRK_CRS_RENEWAL")) Then
                If rsJobHis("TW_TRK_CRS_RENEWAL") <> 0 Then
                    get_Job_History_Data = "P"
                Else
                    get_Job_History_Data = rsJobHis(xAsField)
                End If
            Else
                get_Job_History_Data = rsJobHis(xAsField)
            End If
        End If
    End If
    rsJobHis.Close
    Set rsJobHis = Nothing
End Function
