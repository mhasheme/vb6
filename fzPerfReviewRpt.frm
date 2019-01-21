VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRPerfReviewRpt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Peformance Review"
   ClientHeight    =   10950
   ClientLeft      =   435
   ClientTop       =   870
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   11400
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   10695
      LargeChange     =   300
      Left            =   11760
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   58
      Top             =   120
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   10695
      Left            =   120
      TabIndex        =   34
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
         Height          =   8895
         Left            =   0
         ScaleHeight     =   8865
         ScaleWidth      =   11385
         TabIndex        =   35
         Top             =   0
         Width           =   11415
         Begin VB.ComboBox comFTMonth 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "fzPerfReviewRpt.frx":0000
            Left            =   2050
            List            =   "fzPerfReviewRpt.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "Select Full Time Month"
            Top             =   5230
            Width           =   1590
         End
         Begin VB.DriveListBox Drive1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6195
            TabIndex        =   62
            Top             =   5760
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton cmdLocation 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9240
            TabIndex        =   23
            Tag             =   "Click to select the location"
            Top             =   5760
            Width           =   375
         End
         Begin VB.TextBox txtFilePath 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2050
            TabIndex        =   21
            Tag             =   "Path to save the form"
            Top             =   5760
            Width           =   4050
         End
         Begin VB.DirListBox Dir1 
            BackColor       =   &H00FFFFFF&
            Height          =   2565
            Left            =   6195
            TabIndex        =   22
            Tag             =   "Path"
            Top             =   6120
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CheckBox chkShowMedical 
            Caption         =   "Show Medical Contacts"
            Height          =   285
            Left            =   7320
            TabIndex        =   31
            Top             =   7200
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CheckBox chkForAudit 
            Caption         =   "For Data Audit"
            Height          =   285
            Left            =   7320
            TabIndex        =   32
            Top             =   7320
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkWeeklyEmpList 
            Caption         =   "Show Weekly Employee List"
            Height          =   285
            Left            =   7320
            TabIndex        =   30
            Top             =   6600
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Tag             =   "Final sorting of records"
            Top             =   7455
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.CheckBox chkLastDay 
            Caption         =   "Show Last Day"
            Height          =   285
            Left            =   7320
            TabIndex        =   29
            Top             =   6360
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "Final sorting of records"
            Top             =   7905
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Tag             =   "First Level of grouping records"
            Top             =   7560
            Visible         =   0   'False
            Width           =   2325
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
            Left            =   9210
            TabIndex        =   27
            Tag             =   "40-Date upto and including this date forward"
            Top             =   7785
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   2
            Left            =   7500
            TabIndex        =   26
            Tag             =   "40-Date from and including this date forward"
            Top             =   7785
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
            Left            =   6840
            TabIndex        =   28
            Tag             =   "00-Benefit - Group Code"
            Top             =   8100
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
         Begin INFOHR_Controls.DateLookup dlpReviewFrom 
            Height          =   285
            Left            =   1740
            TabIndex        =   18
            Tag             =   "40-Performance Review From"
            Top             =   4875
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1240
         End
         Begin INFOHR_Controls.DateLookup dlpReviewTo 
            Height          =   285
            Left            =   3840
            TabIndex        =   19
            Tag             =   "40-Performance Review To"
            Top             =   4875
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1240
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Full Time Month"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   5290
            Width           =   1125
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   63
            Top             =   4920
            Width           =   195
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Path to save the file to:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   5805
            Width           =   1635
         End
         Begin VB.Label lblAsOf 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Perf. Review From"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   4920
            Width           =   1305
         End
         Begin VB.Label lblBenGroup 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5760
            TabIndex        =   59
            Top             =   8100
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblGrp 
            BackStyle       =   0  'Transparent
            Caption         =   "Work History Sort"
            Height          =   375
            Index           =   1
            Left            =   5760
            TabIndex        =   57
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
            TabIndex        =   56
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
            Left            =   5760
            TabIndex        =   55
            Top             =   7080
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   4560
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
            TabIndex        =   52
            Top             =   1890
            Width           =   630
         End
         Begin VB.Label lblRep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   4200
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
            TabIndex        =   50
            Top             =   3870
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
            TabIndex        =   49
            Top             =   3540
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
            TabIndex        =   48
            Top             =   3210
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
            TabIndex        =   47
            Top             =   900
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
            Left            =   480
            TabIndex        =   46
            Top             =   7905
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
            Left            =   480
            TabIndex        =   45
            Top             =   7590
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
            Left            =   360
            TabIndex        =   44
            Top             =   7320
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
            TabIndex        =   43
            Top             =   0
            Width           =   1575
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
            TabIndex        =   42
            Top             =   2880
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
            TabIndex        =   41
            Top             =   2550
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
            TabIndex        =   40
            Top             =   2220
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
            TabIndex        =   39
            Top             =   1560
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
            TabIndex        =   38
            Top             =   1230
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
            TabIndex        =   37
            Top             =   570
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
            TabIndex        =   36
            Top             =   240
            Width           =   555
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
   End
End
Attribute VB_Name = "frmRPerfReviewRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQLQ
Public wdApp As Object 'As Word.Application
Public wrdDoc As Object
Dim intPrfCatRec As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
    Dim X%
    
    On Error GoTo PrntErr
    
    If CriCheck() Then
        If Not PrtForm("Performance Review Criteria", Me) Then Exit Sub
        
        Call set_PrintState(False)
        
        X% = Cri_SetAll()
        
        'Me.vbxCrystal.Destination = 1
        MDIMain.Timer1.Enabled = False
        'Me.vbxCrystal.Action = 1
        'vbxCrystal.Reset
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
    Dim X%
    Dim strWHand As String
    On Error GoTo CRW_Err
    
    If CriCheck() Then
        Screen.MousePointer = HOURGLASS
        Call set_PrintState(False)
        
        'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
        'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
        Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
        
        X% = Cri_SetAll()
        
        'Me.vbxCrystal.Destination = 0
        MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        'Me.vbxCrystal.Action = 1
        'vbxCrystal.Reset
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

Private Sub cmdLocation_Click()
On Error GoTo Error_Dir
    If Dir1.Visible = True Then
        txtFilePath.Text = Dir1.Path
        If Len(Trim(txtFilePath.Text)) > 0 Then
            If Right(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        End If
        
        cmdLocation.Caption = ">"
        Drive1.Visible = False
        Dir1.Visible = False
        cmdLocation.Left = 6200 '6800
    Else
        cmdLocation.Caption = "<"
        Drive1.Visible = True
        Dir1.Visible = True
        cmdLocation.Left = 9240 '9840
        If Len(Trim(txtFilePath.Text)) > 0 And Left(txtFilePath.Text, 1) <> "\" And Left(txtFilePath.Text, 1) <> "/" Then
            If Dir(txtFilePath.Text, vbDirectory) <> "" Then
                Drive1.Drive = txtFilePath.Text
                Dir1.Path = Drive1.Drive
                Dir1.Path = txtFilePath.Text
            End If
        End If
    End If
    Exit Sub
Error_Dir:
If Err.Number = 52 Then
    Exit Sub
End If
End Sub

Private Sub comFTMonth_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comFTMonthLoad()
    comFTMonth.AddItem "January"
    comFTMonth.AddItem "February"
    comFTMonth.AddItem "March"
    comFTMonth.AddItem "April"
    comFTMonth.AddItem "May"
    comFTMonth.AddItem "June"
    comFTMonth.AddItem "July"
    comFTMonth.AddItem "August"
    comFTMonth.AddItem "September"
    comFTMonth.AddItem "October"
    comFTMonth.AddItem "November"
    comFTMonth.AddItem "December"
    comFTMonth.AddItem ""
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
    
    If glbMulti Then
        EECri = "HR_JOB_HISTORY.JH_ORG IN  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    Else
        EECri = "HREMP.ED_ORG IN  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
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
    
    DivCri = "(HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    Else
        glbstrSelCri = DivCri
    End If
End Sub

Private Sub Cri_FTMonth()
    Dim EECri As String
    
    If Len(comFTMonth.Text) <= 0 Then Exit Sub
    
    EECri = "(Month(HREMP.ED_USRDAT1) = " & Get_MonthNo(comFTMonth.Text) & ")"
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_EE()
    Dim EECri As String
    
    If Len(elpEEID.Text) <= 0 Then Exit Sub
    
    EECri = "HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
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

    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & "HR_JOB_HISTORY.JH_REPTAU = " & Trim(elpRept(0).Text) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND HR_JOB_HISTORY.JH_REPTAU2 = " & Trim(elpRept(1).Text) & " "
        Else
            EECri = EECri & "HR_JOB_HISTORY.JH_REPTAU2 = " & Trim(elpRept(1).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND HR_JOB_HISTORY.JH_REPTAU3 = " & Trim(elpRept(2).Text) & " "
        Else
            EECri = EECri & "HR_JOB_HISTORY.JH_REPTAU3 = " & Trim(elpRept(2).Text) & " "
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
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "(HR_JOB_HISTORY.JH_SDATE >= " & Date_SQL(dlpDateRange(0).Text)
        TempCri = TempCri & " AND HR_JOB_HISTORY.JH_SDATE <= " & Date_SQL(dlpDateRange(1).Text) & ")"
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "(HR_JOB_HISTORY.JH_SDATE >= " & Date_SQL(dlpDateRange(0).Text) & ")"
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "(HR_JOB_HISTORY.JH_SDATE <= " & Date_SQL(dlpDateRange(1).Text) & ")"
        GoTo Cri_FTDatst
    End If

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
    
    If Len(clpJob.Text) <= 0 Then Exit Sub
        
    PosCri = "(HR_JOB_HISTORY.JH_JOB = '" & clpJob.Text & "')"
    
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
        
    GirdCri = "(HR_JOB_HISTORY.JH_GRID = '" & clpGrid.Text & "')"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & GirdCri
    Else
        glbstrSelCri = GirdCri
    End If
End Sub

Private Sub Cri_PT()
    Dim EECri As String
    
    If Len(clpPT.Text) < 1 Then Exit Sub
    
    If glbMulti Then
        EECri = "(HR_JOB_HISTORY.JH_PT in ('" & Replace(clpPT.Text, ",", "','") & "'))"
    Else
        EECri = "(HREMP.ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "'))"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_BenefitGroup()
    Dim EECri As String
    
    If Len(clpCode(10).Text) < 1 Then Exit Sub
    
    EECri = "(ED_BENEFIT_GROUP = '" & clpCode(10).Text & "')"
    
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End Sub

Private Function Cri_SetAll()
    Dim strRName$
    Dim intGenForm As Integer
    
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
    Call Cri_FTMonth
    
    
    'Call function to generate the MS Word Forms using the templates provided
    intPrfCatRec = 0
    intGenForm = Generate_Word_Forms()
    If intGenForm > 0 Then
        'Successfully generated forms
        MsgBox "Employee Performance Review Report(s) generated successfully.", vbOKOnly + vbInformation, "Performance Review Report"
    ElseIf intGenForm = 0 Then
        'Employees not found
        MsgBox "No Employees in this selection criteria.", vbOKOnly + vbInformation, "Performance Review Report"
    Else
        'Form generation was unsuccessful
        MsgBox "Problem generating Performance Review Report.", vbOKOnly + vbCritical, "Performance Review Report"
    End If
    
    Cri_SetAll = True

    Screen.MousePointer = DEFAULT

Exit Function
modSetCriteria_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select Report Criteria", "Performance Review Report", "Select")
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
    
    Cri_Sorts = 0
    
    'first set primary grouping
    z% = 0
    X% = 0
    
    grpField$ = getEGroup(comGroup(0).Text)
    Y% = X% + 1
    
    If comGroup(0) = "(none)" Then grpField$ = "{@EFullName}"
    Call setRptLabel(Me, 0)
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(X%) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(X%) = grpCond$
    
    Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
    Dim EECri As String, LocCri As String
    
    If Len(clpCode(2).Text) <= 0 Then Exit Sub
    
    If Len(clpCode(2).Text) > 0 Then
        EECri = "(HREMP.ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')) "
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
            CodeCri = "(" & strCd$ & " in  ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
            
            If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
                CodeCri = "(((" & strCd$ & ") = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ((" & strCd$ & ") = 'ALL" & clpCode(intIdx%).Text & "') )"
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
    Dim X%, I
    
    CriCheck = False
    
    If Not clpDiv.ListChecker Then
    'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
        'MsgBox lStr("Invalid Division")
        'clpDiv.SetFocus
        Exit Function
    End If
    
    If Not clpDept.ListChecker Then
    'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
        'MsgBox "Invalid Department"
        'clpDept.SetFocus
        Exit Function
    End If
        
    For X% = 0 To 2
        If Not clpCode(X).ListChecker Then Exit Function
    Next X%
    For X% = 7 To 9
        If Not clpCode(X).ListChecker Then Exit Function
    Next X%
    
    If Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
        MsgBox "Invalid Job Code"
        clpJob.SetFocus
        Exit Function
    End If
    
    If Not clpPT.ListChecker Then
    'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
        'MsgBox lStr("Category code must be valid")
        'clpPT.SetFocus
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
       
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
            MsgBox "To Date cannot be prior to From Date!"                       '
            Me.dlpDateRange(0).SetFocus                                         '
            Exit Function                                                       '
        End If
    End If
    
    For I = 0 To 2
        If elpRept(I).Caption = "Enter Valid Employee #" Then
            MsgBox "Invalid Reporting Authority"
            elpRept(I).SetFocus
            Exit Function
        End If
    Next
    
    If Not elpEEID.ListChecker Then
        Exit Function
    End If
    
    If Not IsDate(dlpReviewFrom.Text) Then
        MsgBox "Invalid Performance Review From Date"
        dlpReviewFrom.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpReviewTo.Text) Then
        MsgBox "Invalid Performance Review To Date"
        dlpReviewTo.SetFocus
        Exit Function
    End If
       
    If IsDate(dlpReviewFrom.Text) And IsDate(dlpReviewTo.Text) Then
        If DaysBetween(dlpReviewFrom.Text, dlpReviewTo.Text) < 0 Then
            MsgBox "Performance Review To Date cannot be prior to From Date"
            Me.dlpReviewFrom.SetFocus
            Exit Function
        End If
    End If
    
    If Len(Trim(txtFilePath.Text)) > 0 Then
        If Right(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
    Else
        MsgBox "'Path to save the file to' cannot be blank"
        txtFilePath.SetFocus
        Exit Function
    End If
    If Dir(txtFilePath.Text, vbDirectory) = "" Then
        MsgBox "Invalid File Path:" & Chr(10) & "[" & txtFilePath.Text & "]"
        txtFilePath.SetFocus
        Exit Function
    End If
    
    CriCheck = True
    
End Function

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    On Error GoTo Load_Error
    
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
    Call comFTMonthLoad
    Call comGrpLoad
    
    'Ticket #17029
    If gsFRIESENSWORDPATH Then
        txtFilePath.Text = GetComPreferEmail("FRIESENSWORDPATH")
    End If
    If Len(txtFilePath) = 0 Then
        txtFilePath.Text = glbIHRREPORTS
    End If
    Drive1.Drive = txtFilePath.Text
    Dir1.Path = Drive1.Drive
    Dir1.Path = txtFilePath.Text
    
    'txtFilePath.Text = glbIHRREPORTS
    cmdLocation.Left = 6200 '6800
    Dir1.Visible = False
    
'    If Me.Caption = "Employee Profile Report" Then
'        lblGrp(1).Visible = True
'        comGroup(2).Visible = True
'    Else
'        lblGrp(1).Visible = False
'        comGroup(2).Visible = False
'    End If
    
    If glbLinamar Then clpCode(7).MaxLength = 8
    If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
    If glbCompSerial = "S/N - 2381W" Then clpCode(0).MaxLength = 6
    
    Call INI_Controls(Me)
    
    panDetails.BorderStyle = 0 'no border
    panWindow.BevelOuter = 0 ' no bevel

    Screen.MousePointer = DEFAULT
    
    Exit Sub
    
Load_Error:
    If Err.Number = 76 Then
        MsgBox "Invalid 'Path to save the file to:'", vbOKOnly, "Invalid Path"
        Resume Next
    Else
        Resume Next
    End If
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form Resize", "Performance Review Report", "Form Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmRPerfReviewRpt = Nothing  'carmen apr 2000
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
        
    EECri = "(HREMP.ED_SHIFT= '" & txtShift.Text & "')"

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
    Dim dtYYY%, dtMM%, dtDD%, X%
    Dim FromDate, ToDate, SQLQ
    Dim RsHRPARCO As New ADODB.Recordset
    
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        TempCri = "(HREMP.ED_SFDATE >= " & Date_SQL(dlpDateRange(2).Text) & ") and "
        TempCri = TempCri & " (HREMP.ED_STDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") "
        GoTo Cri_FTDatst
    End If
    
    If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
        If Len(dlpDateRange(2).Text) > 0 Then
            TempCri = "(HREMP.ED_SFDATE  >= " & Date_SQL(dlpDateRange(2).Text) & ")"
            GoTo Cri_FTDatst
        End If
        If Len(dlpDateRange(3).Text) > 0 Then
            TempCri = TempCri & "(HREMP.ED_STDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") "
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
    
    GetJobCodeDesc = xStr
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
    GetTABLDesc = xStr
End Function

Private Function Generate_Word_Forms()
    Dim rsEmp As New ADODB.Recordset
    Dim rsPerf As New ADODB.Recordset
    Dim rsJobDoc As New ADODB.Recordset
    Dim rsHRJob As New ADODB.Recordset
    Dim rsPerJobGrp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEmpNo As String
    Dim xlsFileTmpl, xlsFileMat As String
    Dim xRow, xRecNum, xPrfCatNo, xPrfNo As Integer
    Dim strReptAuth1 As String
    Dim xLstPerfCat As String
    Dim mydoc As Document
    Dim docrange
    Dim xFTDate, xFTAnnStartDt, xFTAnnEndDt
    
    
    On Error GoTo Err_Word_Forms

    Generate_Word_Forms = 0

    Screen.MousePointer = HOURGLASS

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    'Retrieve the records to generate the forms for
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_DEPTNO, ED_DIV, ED_ORG, ED_SFDATE, ED_USRDAT1, "
    SQLQ = SQLQ & " JH_JOB, JH_SDATE, JH_REPTAU FROM HREMP, HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE ED_EMPNBR = JH_EMPNBR AND JH_CURRENT <>0 "
    If Len(glbstrSelCri) > 0 Then
        SQLQ = SQLQ & " AND " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "[", "("), "]", ")")
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    'Get the Word Template
    xlsFileTmpl = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "PERFORMANCE REVIEW.dot"

    If Not rsEmp.EOF Then
        xRecNum = rsEmp.RecordCount
        xRow = 1
        
        'Create new hidden instance of Word.
        Set wdApp = CreateObject("Word.Application")

        Do While Not rsEmp.EOF
            MDIMain.panHelp(0).FloodPercent = (xRow / xRecNum) * 100

'            'Retrieve Staff Profile records
'            SQLQ = "SELECT PH_EMPNBR,PH_PREVIEW,PH_CATECODE,PH_COMMENTS FROM HR_PERFORM_FRIESEN "
'            SQLQ = SQLQ & " WHERE PH_EMPNBR = " & rsEmp("ED_EMPNBR")
'            SQLQ = SQLQ & " AND (PH_PREVIEW >= " & Date_SQL(dlpReviewFrom.Text)
'            SQLQ = SQLQ & " AND PH_PREVIEW <= " & Date_SQL(dlpReviewTo.Text) & ")"
'            SQLQ = SQLQ & " ORDER BY PH_CATECODE,PH_PREVIEW DESC"
'            rsPerf.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            If rsPerf.EOF Then
'                GoTo Next_Employee
'            Else
'                intPrfCatRec = 1
'                rsPerf.MoveFirst
'            End If
            
            With wdApp
                If Dir(xlsFileTmpl) = "" Then
                    MDIMain.panHelp(0).FloodType = 1
                    MDIMain.panHelp(0).FloodPercent = 0
                    MDIMain.panHelp(0).Caption = "Please wait..."
                    Screen.MousePointer = DEFAULT
                    MsgBox "There is no " & xlsFileTmpl
                    intPrfCatRec = 2
                    Generate_Word_Forms = 0
                    Exit Function
                End If

                'Filename for the Word Form to Save As
                xlsFileMat = txtFilePath.Text & IIf(Right(txtFilePath.Text, 1) = "\", "", "\") & "PE_" & rsEmp("ED_SURNAME") & "_" & rsEmp("ED_FNAME") & "_" & Format(Now, "mmddyy") & ".doc"
                
                'Delete the word document if already exists
                If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

                'Set Word object as the template
                Set wrdDoc = .Documents.Add(xlsFileTmpl, False)

                'Show this instance of Word.
                '.Visible = True

                'Make the word doc Active
                .Documents(wrdDoc).Activate

                'Update the bookmark fields in the Word template with database values
                .ActiveDocument.FormFields("txtEmpName").Result = rsEmp("ED_FNAME") & " " & rsEmp("ED_SURNAME")
                .ActiveDocument.FormFields("intEmpNo").Result = rsEmp("ED_EMPNBR")
                If Not IsNull(rsEmp("JH_REPTAU")) Then
                    .ActiveDocument.FormFields("txtRepAuth1").Result = GetEmpData(rsEmp("JH_REPTAU"), "ED_FNAME") & " " & GetEmpData(rsEmp("JH_REPTAU"), "ED_SURNAME")
                Else
                    .ActiveDocument.FormFields("txtRepAuth1").Result = ""
                End If
                .ActiveDocument.FormFields("txtDiv").Result = Get_Division_Name(rsEmp("ED_DIV"))
                .ActiveDocument.FormFields("txtDept").Result = GetDeptName(rsEmp("ED_DEPTNO"), "DF_NAME")
                .ActiveDocument.FormFields("txtPosition").Result = GetJobCodeDesc(rsEmp("JH_JOB"))
                .ActiveDocument.FormFields("dtStatusFromDt").Result = Format(rsEmp("ED_USRDAT1"), "MMMM d, yyyy")
                .ActiveDocument.FormFields("dtFromDate").Result = Format(dlpReviewFrom.Text, "MMMM d, yyyy")
                .ActiveDocument.FormFields("dtToDate").Result = Format(dlpReviewTo.Text, "MMMM d, yyyy")

'                xLstPerfCat = ""
'                xPrfCatNo = 0
'                xPrfNo = 0
'                Do While Not rsPerf.EOF
'                    If rsPerf("PH_CATECODE") <> xLstPerfCat Then
'                        'New Section of Performance Category
'                        xPrfCatNo = xPrfCatNo + 1   '# of Performance Category
'                        xPrfNo = 0                  'restart Perf. Comments #
'                        xLstPerfCat = rsPerf("PH_CATECODE")
'
'                        If xPrfCatNo <= 18 Then
'                            .ActiveDocument.FormFields("txtPerfCateg" & xPrfCatNo).Result = GetTABLDesc("SDPG", rsPerf("PH_CATECODE"))
'                        Else
'                            'Maximum # of Performance Category bookmarks in template reached. Skip now.
'                        End If
'                    End If
'
'                    xPrfNo = xPrfNo + 1     'Bookmark # for Comments
'
'                    If xPrfNo <= 10 And xPrfCatNo <= 18 Then
'                        .ActiveDocument.FormFields("txtPerf" & xPrfCatNo & xPrfNo).Result = Format(rsPerf("PH_PREVIEW"), "mm/dd/yyyy")
'                        .ActiveDocument.FormFields("txtPerfComm" & xPrfCatNo & xPrfNo).Result = IIf(IsNull(rsPerf("PH_COMMENTS")), "", rsPerf("PH_COMMENTS"))
'                    Else
'                        'Maximum # of Comments or Performance Category bookmarks in template reached. Skip comments now.
'                    End If
'
'                    rsPerf.MoveNext
'                Loop
                
                'For each Performance Categories check if there are any Staff Profile entries for the
                'date range specified by the user as the selection criteria or by FT Date if available
                xFTDate = GetEmpData(rsEmp("ED_EMPNBR"), "ED_USRDAT1")
                If xFTDate <> "" And Not IsNull(xFTDate) Then
                    xFTAnnEndDt = DateAdd("d", -1, CVDate(xFTDate))
                    'Does the year change when deducting a day from the FT Date?
                    If Year(xFTDate) - Year(xFTAnnEndDt) > 0 Then
                        'Ticket #28209 - Check for Leap Year - giving Type mismatch error
                        If month(xFTAnnEndDt) = 2 And Day(xFTAnnEndDt) = 29 And GetLeapYear(Year(Now)) Then
                            'Last year was not Leap Year so change the End Date to 28th otherwise -1 year will give an error
                            xFTAnnEndDt = DateAdd("d", -1, CVDate(xFTAnnEndDt))
                        End If
                        xFTAnnEndDt = CVDate(month(xFTAnnEndDt) & "/" & Day(xFTAnnEndDt) & "/" & Year(Now) - 1)
                        
                        xFTAnnStartDt = DateAdd("d", 1, CVDate(xFTAnnEndDt))
                        'Ticket #28209 - Check for Leap Year - giving Type mismatch error
                        If month(xFTAnnStartDt) = 2 And Day(xFTAnnStartDt) = 29 And GetLeapYear(Year(Now)) Then
                            'Last Year was not Leap Year so change the Start Date to 28th otherwise -1 year will give an error
                            xFTAnnStartDt = DateAdd("d", -1, CVDate(xFTAnnStartDt))
                        End If
                        xFTAnnStartDt = CVDate(month(xFTAnnStartDt) & "/" & Day(xFTAnnStartDt) & "/" & Year(Now) - 1)
                    Else
                        'Ticket #28209 - Check for Leap Year - giving Type mismatch error
                        If month(xFTAnnEndDt) = 2 And Day(xFTAnnEndDt) = 29 And Not GetLeapYear(Year(Now)) Then
                            'Current Year is not Leap Year change the End Date to 28th otherwise changing date to Current will give an error
                            xFTAnnEndDt = DateAdd("d", -1, CVDate(xFTAnnEndDt))
                        End If
                        xFTAnnEndDt = CVDate(month(xFTAnnEndDt) & "/" & Day(xFTAnnEndDt) & "/" & Year(Now))
                        
                        xFTAnnStartDt = DateAdd("d", 1, CVDate(xFTAnnEndDt))
                        'Ticket #28209 - Check for Leap Year - giving Type mismatch error
                        If month(xFTAnnStartDt) = 2 And Day(xFTAnnStartDt) = 29 And GetLeapYear(Year(Now)) Then
                            'Last Year was not Leap Year so change the Start Date to 28th otherwise -1 year will give an error
                            xFTAnnStartDt = DateAdd("d", -1, CVDate(xFTAnnStartDt))
                        End If
                        xFTAnnStartDt = CVDate(month(xFTAnnStartDt) & "/" & Day(xFTAnnStartDt) & "/" & Year(Now) - 1)
                    End If
                    'Review Period
                    .ActiveDocument.FormFields("dtFromDate").Result = Format(xFTAnnStartDt, "MMMM d, yyyy")
                    .ActiveDocument.FormFields("dtToDate").Result = Format(xFTAnnEndDt, "MMMM d, yyyy")
                End If
                
                'Get the Position Group
                SQLQ = "SELECT JB_CODE, JB_GRPCD FROM HRJOB WHERE JB_CODE = '" & rsEmp("JH_JOB") & "'"
                rsHRJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHRJob.EOF Then
                    'Initialise
                    xPrfCatNo = 0
                    
                    'Get the list of Performance Categories linked to the Position Group
                    SQLQ = "SELECT PJ_GRPCD, PJ_CATECODE FROM HR_PERF_JOBGRP "
                    SQLQ = SQLQ & " WHERE PJ_GRPCD = '" & rsHRJob("JB_GRPCD") & "'"
                    SQLQ = SQLQ & " AND PJ_FRIESENS_FORMS <> 0"
                    SQLQ = SQLQ & " ORDER BY PJ_CATECODE"
                    rsPerJobGrp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsPerJobGrp.EOF Then
                        rsPerJobGrp.MoveFirst
                        
                        Do While Not rsPerJobGrp.EOF
                            'New Section - Print the Performance Category in the form as the section heading
                            xPrfCatNo = xPrfCatNo + 1   '# of Performance Category
                            xPrfNo = 0                  'restart Staff Profile Comments #
                            
                            If xPrfCatNo <= 18 Then
                                .ActiveDocument.FormFields("txtPerfCateg" & xPrfCatNo).Result = GetTABLDesc("SDPG", rsPerJobGrp("PJ_CATECODE"))
                            Else
                                'Maximum # of Performance Category bookmarks in the template is reached. Skip now.
                                GoTo Get_Expectation_Document
                            End If
                                                        
                            'Retrieve Staff Profile records
                            SQLQ = "SELECT PH_EMPNBR,PH_PREVIEW,PH_CATECODE,PH_COMMENTS FROM HR_PERFORM_FRIESEN "
                            SQLQ = SQLQ & " WHERE PH_EMPNBR = " & rsEmp("ED_EMPNBR")
                            If xFTDate <> "" And Not IsNull(xFTDate) Then
                                SQLQ = SQLQ & " AND (PH_PREVIEW >= " & Date_SQL(xFTAnnStartDt)
                                SQLQ = SQLQ & " AND PH_PREVIEW <= " & Date_SQL(xFTAnnEndDt) & ")"
                            Else
                                SQLQ = SQLQ & " AND (PH_PREVIEW >= " & Date_SQL(dlpReviewFrom.Text)
                                SQLQ = SQLQ & " AND PH_PREVIEW <= " & Date_SQL(dlpReviewTo.Text) & ")"
                            End If
                            SQLQ = SQLQ & " AND PH_CATECODE = '" & rsPerJobGrp("PJ_CATECODE") & "'"
                            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
                            rsPerf.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsPerf.EOF Then
                                'Staff Profile entries found
                                rsPerf.MoveFirst
                                
                                'Print all the Staff Profile entries for this Perf Category in the form
                                Do While Not rsPerf.EOF
                                    xPrfNo = xPrfNo + 1     'Bookmark # for Comments
                                    
                                    If xPrfNo <= 10 And xPrfCatNo <= 18 Then
                                        .ActiveDocument.FormFields("txtPerf" & xPrfCatNo & xPrfNo).Result = Format(rsPerf("PH_PREVIEW"), "mm/dd/yyyy")
                                        'Getting "String too long" - error 4609 - workaround is the line below
                                        'it's a bug in MS Word since Word 2000
                                        '.ActiveDocument.FormFields("txtPerfComm" & xPrfCatNo & xPrfNo).Result = IIf(IsNull(rsPerf("PH_COMMENTS")), "", rsPerf("PH_COMMENTS"))
                                        '.ActiveDocument.Bookmarks("txtPerfComm" & xPrfCatNo & xPrfNo).Range.Fields(1).Result.Text = IIf(IsNull(rsPerf("PH_COMMENTS")), "", rsPerf("PH_COMMENTS"))
                                        '.ActiveDocument.Bookmarks("txtPerfComm" & xPrfCatNo & xPrfNo).Range.Fields(1).Result.Text = IIf(IsNull(rsPerf("PH_COMMENTS")), "", Replace(Replace(rsPerf("PH_COMMENTS"), vbLf, ""), vbCr, ""))
                                        If Not IsNull(rsPerf("PH_COMMENTS")) Then
                                            .ActiveDocument.Bookmarks("txtPerfComm" & xPrfCatNo & xPrfNo).Range.Fields(1).Result.Text = IIf(IsNull(rsPerf("PH_COMMENTS")), "", Replace(rsPerf("PH_COMMENTS"), vbCrLf, " "))
                                        End If
                                    Else
                                        'Maximum # of Comments or Performance Category bookmarks in the template is reached. Skip comments now.
                                        GoTo Next_Perf_Category
                                    End If
                                
                                    rsPerf.MoveNext
                                Loop
                            End If
Next_Perf_Category:
                            rsPerf.Close
                            Set rsPerf = Nothing
                
                            rsPerJobGrp.MoveNext
                        Loop
                    End If
Get_Expectation_Document:
                    rsPerJobGrp.Close
                    Set rsPerJobGrp = Nothing
                End If

                rsHRJob.Close
                Set rsHRJob = Nothing
                
                'Get the filename and location for the Expectation Document of the employee's
                'Position to append to the Performance Review document.
                SQLQ = "SELECT JD_JOB,JD_DOC_TYPE,JD_FILE_LINK FROM HR_JOB_DOCUMENT"
                SQLQ = SQLQ & " WHERE JD_JOB = '" & rsEmp("JH_JOB") & "'"
                SQLQ = SQLQ & " AND JD_DOC_TYPE = 'PEEX'"
                rsJobDoc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsJobDoc.EOF Then
                    'Expectation Document found for this employee's Job
                    'Append the Expectation Document of this employee's Position to this Performance Review doc.
                    .ActiveDocument.Unprotect
                    Set docrange = .ActiveDocument.Range    'can use Content instead of Range
                    docrange.Collapse wdCollapseEnd
                    docrange.InsertBreak wdSectionBreakNextPage
                    Set docrange = .ActiveDocument.Range
                    docrange.Collapse wdCollapseEnd
                    
                    docrange.InsertFile (rsJobDoc("JD_FILE_LINK"))
                    'docrange.InsertFile "c:\Report template programming details.doc"
                    
                    .ActiveDocument.Protect wdAllowOnlyFormFields, NoReset:=True
                    
                Else
                    'Expectation Document not found for this employee's Job
                    MsgBox "Expectation Document not found for Job: " & GetJobCodeDesc(rsEmp("JH_JOB")), vbOKOnly + vbInformation, "Expectation Document"
                End If
                rsJobDoc.Close
                Set rsJobDoc = Nothing
                
                'Set docrange = ActiveDocument.Range
                '.ActiveDocument.Range.Collapse wdCollapseEnd
                '.ActiveDocument.Range.InsertBreak wdSectionBreakNextPage
                'Set docrange = ActiveDocument.Range
                '.ActiveDocument.Range.Collapse wdCollapseEnd
                '.ActiveDocument.Range.InsertFile "c:\Report template programming details.doc"
                'wdApp.ActiveDocument.Protect wdAllowOnlyFormFields, NoReset
                
                
                'Save the template as the Word Document - with the filename generated above
                wrdDoc.SaveAs xlsFileMat
                
                .ActiveDocument.Close
            End With
            
            Set wrdDoc = Nothing
            
Next_Employee:
            xRow = xRow + 1
            rsEmp.MoveNext
        Loop
        
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit
        Set wdApp = Nothing
        
        Generate_Word_Forms = xRecNum
    Else
        Generate_Word_Forms = 0
    End If
    rsEmp.Close

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
Exit Function

Err_Word_Forms:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    If Err = 1004 Then
        Resume Next
    End If
    If Err = 5174 Then
        MsgBox Err.Description & Chr(10), vbExclamation + vbOKOnly, "Expectation Document not found"
        Resume Next
    End If
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close all reports and try again.", vbExclamation + vbOKOnly, "Error generating report"
        GoTo close_all
    End If
    If Err = 70 Then
        MsgBox Err.Description & Chr(10) & "Please close all reports and try again.", vbExclamation + vbOKOnly, "Error generating report"
        GoTo close_all
    End If
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Word Forms", "Performance Review Report", "Generate Word Form")
    
    wrdDoc.Close wdDoNotSaveChanges
    
close_all:
    wdApp.NormalTemplate.Saved = True
    wdApp.Quit
    Set wdApp = Nothing
    
    Generate_Word_Forms = -1
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
End Function

Public Function Get_MonthNo(xMonthName)
    Select Case UCase(xMonthName)
        Case "JANUARY"
            Get_MonthNo = 1
        Case "FEBRUARY"
            Get_MonthNo = 2
        Case "MARCH"
            Get_MonthNo = 3
        Case "APRIL"
            Get_MonthNo = 4
        Case "MAY"
            Get_MonthNo = 5
        Case "JUNE"
            Get_MonthNo = 6
        Case "JULY"
            Get_MonthNo = 7
        Case "AUGUST"
            Get_MonthNo = 8
        Case "SEPTEMBER"
            Get_MonthNo = 9
        Case "OCTOBER"
            Get_MonthNo = 10
        Case "NOVEMBER"
            Get_MonthNo = 11
        Case "DECEMBER"
            Get_MonthNo = 12
    End Select
End Function
