VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRAssoc 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   11190
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
   ScaleHeight     =   10950
   ScaleWidth      =   11190
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel panWindow 
      Height          =   10095
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   17806
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9735
         Left            =   120
         ScaleHeight     =   9735
         ScaleWidth      =   10095
         TabIndex        =   33
         Top             =   120
         Width           =   10095
         Begin VB.ComboBox cmdComp 
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
            ItemData        =   "fzassoc.frx":0000
            Left            =   2115
            List            =   "fzassoc.frx":0007
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "Select 'Y' for Yes and 'N' for No"
            Top             =   5535
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtShift 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7320
            MaxLength       =   4
            TabIndex        =   14
            Tag             =   "00-Employee Position Shift"
            Top             =   4170
            Visible         =   0   'False
            Width           =   450
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
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Final sorting of records - no totals"
            Top             =   9135
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
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "Third level of grouping records"
            Top             =   8820
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
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "Second level of grouping records"
            Top             =   8505
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
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Tag             =   "First Level of grouping records"
            Top             =   8190
            Width           =   2325
         End
         Begin VB.ComboBox cmdCostOfEmp 
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
            ItemData        =   "fzassoc.frx":0014
            Left            =   2115
            List            =   "fzassoc.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   5160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbCompanyPd 
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
            ItemData        =   "fzassoc.frx":0018
            Left            =   2115
            List            =   "fzassoc.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Tag             =   "Select 'Y' for Yes and 'N' for No"
            Top             =   5880
            Visible         =   0   'False
            Width           =   615
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Tag             =   "10-Enter Employee Number"
            Top             =   2190
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            ShowUnassigned  =   1
            TextBoxWidth    =   7435
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   9
            Tag             =   "00-Association Codes"
            Top             =   2850
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "TDCD"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   4
            Tag             =   "00-Enter Status Code"
            Top             =   1530
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDEM"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpPT 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Tag             =   "EDPT-Category"
            Top             =   1860
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDPT"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   3
            Tag             =   "00-Enter Union Code"
            Top             =   1200
            Width           =   7755
            _ExtentX        =   13679
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
            Left            =   1800
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
            Left            =   1800
            TabIndex        =   1
            Tag             =   "00-Specific Department Desired"
            Top             =   540
            Width           =   7755
            _ExtentX        =   13679
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
            Top             =   210
            Width           =   7755
            _ExtentX        =   13679
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
            Index           =   10
            Left            =   1800
            TabIndex        =   11
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
            Index           =   11
            Left            =   1800
            TabIndex        =   13
            Tag             =   "00-Enter Section Code"
            Top             =   4170
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   9
            Left            =   1800
            TabIndex        =   10
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
            Left            =   3690
            TabIndex        =   8
            Tag             =   "40-Date upto and including this date forward"
            Top             =   2520
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   7
            Tag             =   "40-Date from and including this date forward"
            Top             =   2520
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   34
            Tag             =   "00-Enter Degree Completed Code"
            Top             =   2520
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EUMJ"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   5
            Left            =   1800
            TabIndex        =   12
            Top             =   3840
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EUMJ"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   12
            Left            =   1800
            TabIndex        =   35
            Tag             =   "00-Enter Major Study Code"
            Top             =   2850
            Visible         =   0   'False
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EUMJ"
         End
         Begin Threed.SSCheck chkTerm 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Tag             =   "Check for Include Terminations"
            Top             =   7350
            Visible         =   0   'False
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Include Terminations          "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin INFOHR_Controls.CodeLookup clpJOB 
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Tag             =   "00-Enter Position Code"
            Top             =   4500
            Visible         =   0   'False
            Width           =   7755
            _ExtentX        =   13679
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
            Left            =   1800
            TabIndex        =   16
            Tag             =   "01-School - Code"
            Top             =   4815
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EUSC"
         End
         Begin INFOHR_Controls.CodeLookup clpProv 
            Height          =   285
            Left            =   1800
            TabIndex        =   20
            Tag             =   "31-Province of Residence - Code"
            Top             =   6240
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
            Top             =   6570
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   3
            Left            =   3690
            TabIndex        =   23
            Tag             =   "40-Date upto and including this date forward"
            Top             =   6900
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
            Left            =   1800
            TabIndex        =   22
            Tag             =   "40-Date from and including this date forward"
            Top             =   6900
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin Threed.SSCheck chkDolEntDtls 
            Height          =   255
            Left            =   5280
            TabIndex        =   26
            Tag             =   "Check for Include Terminations"
            Top             =   7350
            Visible         =   0   'False
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Show Actual Amount Details          "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSCheck chkDolEntComment 
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Tag             =   "Check for Include Terminations"
            Top             =   7350
            Visible         =   0   'False
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Show Comments          "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label lblComp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Completed"
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
            Top             =   5595
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblSchool 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "School"
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
            Top             =   4860
            Visible         =   0   'False
            Width           =   600
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   4545
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblFormalDate 
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
            Left            =   5520
            TabIndex        =   61
            Top             =   2550
            Visible         =   0   'False
            Width           =   870
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
            Left            =   6600
            TabIndex        =   60
            Top             =   4215
            Visible         =   0   'False
            Width           =   765
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
            TabIndex        =   59
            Top             =   1905
            Width           =   630
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
            TabIndex        =   58
            Top             =   4215
            Width           =   540
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   3555
            Width           =   1125
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
            TabIndex        =   54
            Top             =   9165
            Width           =   660
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
            Left            =   120
            TabIndex        =   53
            Top             =   8850
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
            Left            =   120
            TabIndex        =   52
            Top             =   8535
            Width           =   885
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
            TabIndex        =   51
            Top             =   8220
            Width           =   885
         End
         Begin VB.Label lblReportGrp 
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
            TabIndex        =   50
            Top             =   7860
            Width           =   1575
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Paid"
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
            Top             =   3885
            Width           =   1020
         End
         Begin VB.Label lblBCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Associations"
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
            Top             =   2895
            Width           =   1365
         End
         Begin VB.Label lblRenewal 
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
            TabIndex        =   47
            Top             =   2565
            Width           =   870
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
            TabIndex        =   46
            Top             =   2235
            Width           =   1290
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
            TabIndex        =   45
            Top             =   1575
            Width           =   450
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
            TabIndex        =   44
            Top             =   1245
            Width           =   420
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
            TabIndex        =   43
            Top             =   585
            Width           =   825
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
            TabIndex        =   42
            Top             =   255
            Width           =   555
         End
         Begin VB.Label lblSelectCrit 
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
            TabIndex        =   41
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label lblCostOfEmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cost of Employment "
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
            Top             =   5220
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label lblCompanyPd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Paid"
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
            Top             =   5940
            Visible         =   0   'False
            Width           =   1020
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
            TabIndex        =   38
            Top             =   6615
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
            TabIndex        =   37
            Top             =   6285
            Width           =   1365
         End
         Begin VB.Label lblDOH 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Original Hire"
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
            Top             =   6945
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   9495
      LargeChange     =   300
      Left            =   10800
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   31
      Top             =   120
      Width           =   300
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   5280
      Top             =   10320
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
Attribute VB_Name = "frmRAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EECri As String, OneSet%, X%

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

  If CriCheck() Then
    If FormAssoc% = True Then
        If Not PrtForm(lStr("Associations") & " Report Criteria", Me) Then Exit Sub
    ElseIf FormDoll% = True Then  'laura Oct 28, 1997
        If Not PrtForm("Dollar Entitlement Report Criteria", Me) Then Exit Sub
    ElseIf FormEduc% = True Then
        If Not PrtForm("Formal Education Report Criteria", Me) Then Exit Sub
    ElseIf FormOther% = True Then
        If Not PrtForm("Other Earnings Report Criteria", Me) Then Exit Sub
    Else
    End If           '~~~~~~Laura Oct 28, 1997
    
    Call set_PrintState(False)
    
    X% = Cri_SetAll()
    
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    MDIMain.Timer1.Enabled = True
    
    Call set_PrintState(True)
  End If
'End If
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
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString  'laura nov 21, 1997

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub chkTerm_Click(Value As Integer)
If chkTerm.Value = True Then
        comGroup(2).Enabled = False
  End If
End Sub

Private Sub cmbCompanyPd_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdComp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()

'For X% = 0 To 3

If frmRAssoc.Caption = lStr("Associations") & " Report" Then                       'Laura Oct 20, 1997
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem lStr("Associations")
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem lStr("Associations")
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Renewal Date"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0

ElseIf frmRAssoc.Caption = "Dollar Entitlement Report" Then    'Laura Oct 20, 1997
    
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
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Entitlement Type"
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Entitlement Type"
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Dates From/To"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False
ElseIf frmRAssoc.Caption = "Formal Education Report" Then
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Degree"
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Degree"
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Month/Year"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False

ElseIf frmRAssoc.Caption = "Other Earnings Report" Then
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem lStr("Region")
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Type"
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Type"
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Dates From/To"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False


End If  'Laura 20 Oct, 1997

End Sub

Private Sub Cri_Attend()
Dim EECri As String, OneSet%, X%
Dim strC2, strCx As String
Dim strCa$

'OneSet% = False

If Len(clpCode(1).Text) = 0 Then Exit Sub

'If OneSet% = 0 Then Exit Sub
If FormAssoc% = True Then
    strCa$ = "HRTRADE.TD_CODE"
ElseIf FormDoll% = True Then
    If chkTerm = True Then
        strCa$ = "HRDOLENT.ED_TYPE"
    Else
        strCa$ = "HRDOLENT.DE_TYPE"
    End If
ElseIf FormOther% = True Then
    strCa$ = "HREARN.EARN_TYPE"
End If

If Len(clpCode(1).Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} in ['" & Replace(clpCode(1).Text, ",", "','") & "'])"
End If

If glbiOneWhere Then
   glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
Else
    glbstrSelCri = EECri
End If

glbiOneWhere = True

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If FormAssoc% = True Then
    If Len(clpCode(intIdx%).Text) > 0 Then
        Select Case intIdx%
            Case 0: strCd$ = "HREMP.ED_LOC"
            Case 1: strCd$ = "HRTRADE.TD_CODE"
            Case 2: strCd$ = "HREMP.ED_ORG" ' 1-> 2
            Case 9: strCd$ = "HREMP.ED_REGION"
            Case 10: strCd$ = "HREMP.ED_ADMINBY"
            Case 11: strCd$ = "HREMP.ED_SECTION"
            Case 3: strCd$ = "HREMP.ED_EMP"
        End Select
        
    End If
ElseIf FormDoll% = True Then
    If Len(clpCode(intIdx%).Text) > 0 Then
        Select Case intIdx%
            Case 0: strCd$ = "HREMP.ED_LOC"
            Case 4, 5, 6: strCd$ = "HRDOLENT.DE_TYPE"
            Case 2: strCd$ = "HREMP.ED_ORG" ' 1->2 Ticket# 6496
            Case 3: strCd$ = "HREMP.ED_EMP"
            Case 9: strCd$ = "HREMP.ED_REGION"
            Case 10: strCd$ = "HREMP.ED_ADMINBY"
            Case 11: strCd$ = "HREMP.ED_SECTION"
            'Case Else: strCd$ = "HREMP.ED_EMP"
        End Select
    End If
ElseIf FormEduc% = True Then
    If Len(clpCode(intIdx%).Text) > 0 Then
        Select Case intIdx%
            Case 0: strCd$ = "HREMP.ED_LOC"
            'Hemu - Begin - Case 4 and 5 are not working based on the selection criteria
            '               entered and so separated it out below
            'Case 3, 4, 5, 6: strCd$ = "HREDU.EU_MAJOR"
            'Case 3, 6: strCd$ = "HREDU.EU_MAJOR"
            'Hemu - End
            Case 2: strCd$ = "HREMP.ED_ORG"
            Case 3: strCd$ = "HREMP.ED_EMP"
            'Hemu - Begin - Case 4 and 5 were not picking the right fields above - fixed here
            Case 4: strCd$ = "HREDU.EU_DEGREE"
            Case 5: strCd$ = "HREDU.EU_MINOR"
            Case 6: strCd$ = "HREDU.EU_SCHOOL" 'Ticket #12465
            'New control for Major study selection
            Case 12: strCd$ = "HREDU.EU_MAJOR"
            'Hemu - End
            Case 9: strCd$ = "HREMP.ED_REGION"
            Case 10: strCd$ = "HREMP.ED_ADMINBY"
            Case 11: strCd$ = "HREMP.ED_SECTION"
        End Select
    End If
ElseIf FormOther% = True Then
    If Len(clpCode(intIdx%).Text) > 0 Then
        Select Case intIdx%
            Case 0: strCd$ = "HREMP.ED_LOC"
            'Case 3, 4, 5, 6: strCd$ = "HREARN.EARN_TYPE"
            Case 2: strCd$ = "HREMP.ED_ORG"
            Case 3: strCd$ = "HREMP.ED_EMP"
            Case 9: strCd$ = "HREMP.ED_REGION"
            Case 7: strCd$ = "HREMP.ED_ADMINBY"
            Case 8: strCd$ = "HREMP.ED_SECTION"
            Case Else: strCd$ = "HREMP.ED_EMP"
        End Select
    End If
End If
If Len(clpCode(intIdx%).Text) > 0 Then
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

Private Sub Cri_CompPd()
Dim EECri As String

If Len(cmbCompanyPd.Text) > 0 Then
    If FormEduc% = True Then
        EECri = "({HREDU.EU_COMPANY_PAID} = '" & cmbCompanyPd.Text & "')"
    Else
        EECri = "({HRTRADE.TD_COMPPD} = '" & cmbCompanyPd.Text & "')"
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

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    If frmRAssoc.Caption = lStr("Associations") & " Report" Then  'Laura 20 Oct, 1997
        TempCri = "({HRTRADE.TD_RENEWDT} "
    ElseIf frmRAssoc.Caption = "Dollar Entitlement Report" Then 'Laura
        TempCri = "({HRDOLENT.DE_FDATE} " 'Laura
        If chkTerm = True Then
            TempCri = "({HRDOLENT.ED_FDATE} "
        End If
    ElseIf frmRAssoc.Caption = "Other Earnings Report" Then
        TempCri = "({HREARN.FDATE} "   'Laura changed from TDate to Fdate
    ElseIf frmRAssoc.Caption = "Formal Education Report" Then
        TempCri = "({HREDU.EU_YEAR} "
    End If  'laura

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
          'TempCri = "({HRTRADE.TD_RENEWDT}  "
        If frmRAssoc.Caption = lStr("Associations") & " Report" Then  'Laura 20 Oct, 1997
            TempCri = "({HRTRADE.TD_RENEWDT} "
        ElseIf frmRAssoc.Caption = "Dollar Entitlement Report" Then 'Laura
            If X% = 0 Then
                TempCri = "({HRDOLENT.DE_FDATE} " 'Laura
                If chkTerm = True Then
                    TempCri = "({HRDOLENT.ED_FDATE} "
                End If
            Else
                TempCri = "({HRDOLENT.DE_TDATE} "
                If chkTerm = True Then
                    TempCri = "({HRDOLENT.ED_TDATE} "
                End If
            End If
        ElseIf frmRAssoc.Caption = "Other Earnings Report" Then
            If X% = 0 Then
                TempCri = "({HREARN.FDATE} "   'Laura changed from TDate to Fdate
            Else
                TempCri = "({HREARN.TDATE} "
            End If
        End If  'laura
        
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

Private Sub Cri_DOH()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
    If frmRAssoc.Caption = "Formal Education Report" Then
        TempCri = "({HREMP.ED_DOH} "
    End If  'laura

    dtYYY% = Year(dlpDateRange(2).Text)
    dtMM% = month(dlpDateRange(2).Text)
    dtDD% = Day(dlpDateRange(2).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(3).Text)
    dtMM% = month(dlpDateRange(3).Text)
    dtDD% = Day(dlpDateRange(3).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_DOH
End If

For X% = 2 To 3
    If Len(dlpDateRange(X%).Text) > 0 Then
        If frmRAssoc.Caption = "Formal Education Report" Then
            TempCri = "({HREMP.ED_DOH} "
        End If
        
        If X% = 2 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        
        dtYYY% = Year(dlpDateRange(X%).Text)
        dtMM% = month(dlpDateRange(X%).Text)
        dtDD% = Day(dlpDateRange(X%).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_DOH
    End If
Next X%


Cri_DOH:
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

'Hemu
Private Sub Cri_COE()
    Dim EECri As String
    Dim strCa$
    
    If FormDoll% = True Then
        strCa$ = "HRDOLENT.DE_COST_OF_EMPLOYMENT"
        If chkTerm.Value = True Then
            strCa$ = "HR_DOLENTWRK.ED_COST_OF_EMPLOYMENT"
        End If
    ElseIf FormOther% = True Then
        strCa$ = "HREARN.COST_OF_EMPLOYMENT"
    End If
    
    If (glbOracle) Then
        'Ticket #13188
        If cmdCostOfEmp.Text = "" Then
            EECri = "(1=1)"
        Else
            If cmdCostOfEmp.Text = "Y" Then
                EECri = "{" & strCa$ & "} <> 0"
            End If
            If cmdCostOfEmp.Text = "N" Then
                EECri = "{" & strCa$ & "} = 0"
            End If
        End If
        'If chkCOEFlag Then
        '    EECri = "{" & strCa$ & "} <> 0"
        'Else
        '    EECri = "{" & strCa$ & "} = 0"
        'End If
    Else
        'Ticket #13188
        If cmdCostOfEmp.Text = "" Then
            EECri = "(1=1)"
        Else
            If cmdCostOfEmp.Text = "Y" Then
                EECri = "{" & strCa$ & "}"
            End If
            If cmdCostOfEmp.Text = "N" Then
                EECri = " NOT {" & strCa$ & "}"
            End If
        End If
        'If chkCOEFlag Then
        '    EECri = "{" & strCa$ & "}"
        'Else
        '    EECri = " NOT {" & strCa$ & "}"
        'End If
    End If
    
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True

End Sub

Private Sub SELDOLWRK()

Dim xlen, xxx, xx1
Dim db001 As Database
Dim SQLQ, SQLQ1
Dim xFieldList

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
gdbAdoIhr001.BeginTrans
SQLQ = "DELETE FROM HR_DOLENTWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001W, "HR_DOLENTWRK", "WRKEMP")
SQLQ = "INSERT INTO HR_DOLENTWRK (" & xFieldList & ",WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ1 = Replace(xFieldList, "ED_TYPE_TABL", "DE_TYPE_TABL")
SQLQ1 = Replace(SQLQ1, "ED_TYPE", "DE_TYPE")
SQLQ1 = Replace(SQLQ1, "ED_COST_OF_EMPLOYMENT", "DE_COST_OF_EMPLOYMENT")
SQLQ1 = Replace(SQLQ1, "ED_FDATE", "DE_FDATE")
SQLQ1 = Replace(SQLQ1, "ED_TDATE", "DE_TDATE")
SQLQ1 = Replace(SQLQ1, "ED_ENTITLE", "DE_ENTITLE")
SQLQ1 = Replace(SQLQ1, "ED_ACTUAL", "DE_ACTUAL")
SQLQ1 = Replace(SQLQ1, "ED_REFNBR", "DE_REFNBR")
SQLQ1 = Replace(SQLQ1, "ED_PAIDTO", "DE_PAIDTO")
SQLQ1 = Replace(SQLQ1, "ED_LDATE", "DE_LDATE")
SQLQ1 = Replace(SQLQ1, "ED_LTIME", "DE_LTIME")
SQLQ1 = Replace(SQLQ1, "ED_LUSER", "DE_LUSER")
SQLQ1 = Replace(SQLQ1, "ED_COMMENTS", "DE_COMMENTS")


SQLQ = SQLQ & " SELECT " & Replace(SQLQ1, "ED_PAIDDATE", "DE_PAIDDATE") & ",'" & glbUserID & "' AS WRKEMP "
If Not glbOracle Then
    SQLQ = SQLQ & "FROM HRDOLENT LEFT JOIN HREMP ON HRDOLENT.DE_EMPNBR = HREMP.ED_EMPNBR "
Else
    SQLQ = SQLQ & "FROM HRDOLENT, HREMP WHERE HRDOLENT.DE_EMPNBR = HREMP.ED_EMPNBR(+) "
End If




MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ, , adCmdText
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 50
SQLQ = "INSERT INTO HR_DOLENTWRK (" & xFieldList & ",WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP "


SQLQ1 = Replace(xFieldList, "ED_TYPE_TABL", "DE_TYPE_TABL")
SQLQ1 = Replace(SQLQ1, "ED_TYPE", "DE_TYPE")
SQLQ1 = Replace(SQLQ1, "ED_COST_OF_EMPLOYMENT", "DE_COST_OF_EMPLOYMENT")
SQLQ1 = Replace(SQLQ1, "ED_FDATE", "DE_FDATE")
SQLQ1 = Replace(SQLQ1, "ED_TDATE", "DE_TDATE")
SQLQ1 = Replace(SQLQ1, "ED_ENTITLE", "DE_ENTITLE")
SQLQ1 = Replace(SQLQ1, "ED_ACTUAL", "DE_ACTUAL")
SQLQ1 = Replace(SQLQ1, "ED_REFNBR", "DE_REFNBR")
SQLQ1 = Replace(SQLQ1, "ED_PAIDTO", "DE_PAIDTO")
SQLQ1 = Replace(SQLQ1, "ED_LDATE", "DE_LDATE")
SQLQ1 = Replace(SQLQ1, "ED_LTIME", "DE_LTIME")
SQLQ1 = Replace(SQLQ1, "ED_LUSER", "DE_LUSER")
SQLQ1 = Replace(SQLQ1, "ED_COMMENTS", "DE_COMMENTS")


SQLQ = SQLQ & " SELECT " & Replace(SQLQ1, "ED_PAIDDATE", "DE_PAIDDATE") & ",'" & glbUserID & "' AS WRKEMP "


If Not glbOracle Then
    SQLQ = SQLQ & " FROM Term_DOLENT LEFT JOIN Term_HREMP ON Term_DOLENT.DE_EMPNBR = Term_HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE DE_EMPNBR IS NOT NULL"
Else
    SQLQ = SQLQ & " FROM Term_DOLENT, Term_HREMP WHERE Term_DOLENT.DE_EMPNBR = Term_HREMP.ED_EMPNBR(+) "
    SQLQ = SQLQ & "AND DE_EMPNBR IS NOT NULL"
End If
'SQLQ = SQLQ & "WHERE DE_EMPNBR IS NOT NULL"


MDIMain.panHelp(0).FloodPercent = 60
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)
gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

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
Dim X%, strRName$
Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri

'Call glbCri_Dept(Me)  'laura nov 21, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE

If frmRAssoc.Caption = lStr("Associations") & " Report" Or frmRAssoc.Caption = "Formal Education Report" Then 'Laura oct 20, 1997
  Call Cri_CompPd
End If

'If frmRAssoc.Caption <> "Formal Education Report" Then
  Call Cri_FTDates       'not for education
'End If
Cri_Code (0)
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 9 To 11
    Call Cri_Code(X%)
Next X%

Call Cri_ProvResidence
Call Cri_ProvEmployment

If frmRAssoc.Caption = lStr("Associations") & " Report" Then 'Laura oct 20, 1997
    Call Cri_Attend
    For X% = 2 To 3
        Call Cri_Code(X%)
    Next X%
    ' report name
    If comGroup(0) <> "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzassoc.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzassoc1.rpt"
    End If
    ' set to sorting/grouping criteria
    X% = Cri_Sorts()   ' returns number of sections formated
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If

    Me.vbxCrystal.WindowTitle = lStr("Associations") & " Report"
    Me.vbxCrystal.Connect = RptODBC_SQL
    
ElseIf frmRAssoc.Caption = "Dollar Entitlement Report" Then 'Laura oct 20, 1997
    '~~~~~~~~~copied from FZDOLENT.FRM
    For X% = 2 To 3
        Call Cri_Code(X%)
    Next X%
    Call Cri_Attend
    
    'Hemu
    Call Cri_COE
    'Hemu

    ' report name
    If chkTerm.Value = False Then
        If comGroup(0) <> "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdolent.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdolen1.rpt"
        End If
        X% = Cri_SortsDol()   ' returns number of sections formated
        If chkDolEntComment.Value Then 'Ticket #13188
            Me.vbxCrystal.Formulas(10) = "ShowComments= True"
        End If
        'Ticket #28789 - Show Amount Details
        If chkDolEntDtls.Value Then
            Me.vbxCrystal.Formulas(11) = "ShowDetails= True"
        End If
        
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = glbstrSelCri
        End If
        
        Me.vbxCrystal.WindowTitle = "Dollar Entitlement Report"
        
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Call SELDOLWRK
        If chkDolEntComment.Value Then 'Ticket #13188
            Me.vbxCrystal.Formulas(10) = "ShowComments= True"
        End If
        'Ticket #28789 - Show Amount Details
        If chkDolEntDtls.Value Then
            Me.vbxCrystal.Formulas(11) = "ShowDetails= True"
        End If
        
        If comGroup(0) <> "(none)" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdolentterm.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdolentterm1.rpt"
        End If
       
        X% = Cri_SortsDol()
        glbstrSelCri = Replace(glbstrSelCri, "HREMP", "HR_DOLENTWRK")
        glbstrSelCri = Replace(glbstrSelCri, "HRDOLENT", "HR_DOLENTWRK")
        
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND ({HR_DOLENTWRK.WRKEMP} = '" & glbUserID & "' )"
            Me.vbxCrystal.SelectionFormula = glbstrSelCri
        End If
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
            Me.vbxCrystal.DataFiles(2) = glbIHRDB
            Me.vbxCrystal.DataFiles(3) = glbIHRDB
            Me.vbxCrystal.DataFiles(4) = glbIHRDB
            Me.vbxCrystal.DataFiles(5) = glbIHRDB
            Me.vbxCrystal.DataFiles(6) = glbIHRDB
            Me.vbxCrystal.DataFiles(7) = glbIHRDBW
        End If
        Me.vbxCrystal.WindowTitle = "Dollar Entitlement Report Including Terminations"
    End If
    
ElseIf frmRAssoc.Caption = "Formal Education Report" Then
    For X% = 1 To 5
        Call Cri_Code(X%)
    Next X%
    'Hemu - Begin - New control for the major study selection, earlier selection had problems
    Call Cri_Code(12)
    'Hemu - End
    
    Call Cri_Code(6)        'Ticket #12465 Frank Feb 26, 2007
    Call Cri_EDU_Complete   'Ticket #12465 Frank Feb 26, 2007
    Call Cri_Job            'Ticket #12465 Frank Feb 26, 2007
    
    'Ticket #24906 - Vitalaire Report changes - DOH Date range
    If glbCompSerial = "S/N - 2380W" Then
        Call Cri_DOH
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
    End If
    
    If comGroup(0) <> "(none)" Then
        'Ticket #24906 - Vitalaire Report changes - DOH Date range
        If glbCompSerial = "S/N - 2380W" Then
            strRName$ = glbIHRREPORTS & "SN2380_rzformed.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzformed.rpt"
        End If
    Else
        'Ticket #24906 - Vitalaire Report changes - DOH Date range
        If glbCompSerial = "S/N - 2380W" Then
            strRName$ = glbIHRREPORTS & "SN2380_rzforme1.rpt"
        Else
            strRName$ = glbIHRREPORTS & "rzforme1.rpt"
        End If
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
    
    ' set to sorting/grouping criteria
    X% = Cri_SortsEduc()   ' returns number of sections formated
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    
    Me.vbxCrystal.WindowTitle = "Formal Education Report"
    
    Me.vbxCrystal.Connect = RptODBC_SQL
    
ElseIf frmRAssoc.Caption = "Other Earnings Report" Then
    For X% = 2 To 3
        Call Cri_Code(X%)
    Next X%
    Call Cri_Attend
    
    'Hemu
    Call Cri_COE
    'Hemu
    
    'Ticket #24410 - City of Sarnia - Position Code added
    Call Cri_Job
    
    If comGroup(0) <> "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzoearn.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzoearn1.rpt"
    End If
    X% = Cri_SortsOther()   ' returns number of sections formated
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If

    Me.vbxCrystal.WindowTitle = "Other Earnings Report"
    
    Me.vbxCrystal.Connect = RptODBC_SQL
    
End If  'Laura Oct 20, 1997
'If glbSQL Or glbOracle Then
'    Me.vbxCrystal.Connect = RptODBC_SQL
'End If
Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Assoc ", "Assoc Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_EDU_Complete()
Dim EECri As String, OneSet%, X%

If Len(cmdComp.Text) = 0 Then Exit Sub
EECri = "{HREDU.EU_COMP}= '" & cmdComp.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

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
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0

grpField$ = getEGroup(comGroup(0).Text)

If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    strSFormat$ = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblAssoc.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblAssoc.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(3) & ";" & "{HRTRADE.TD_RENEWDT}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    Else
        grpCond$ = "GROUP" & CStr(1) & ";" & "{@EFullName}" & ";ANYCHANGE;D"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblAssoc.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(3) & ";" & "{HRTRADE.TD_RENEWDT}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblAssoc.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblAssoc.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & "{HRTRADE.TD_RENEWDT}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        End If
     Else
        grpCond$ = "GROUP" & CStr(1) & ";" & "{@EFullName}" & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblAssoc.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & "{HRTRADE.TD_RENEWDT}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        End If
    End If
End If
If frmRAssoc.Caption = lStr("Associations") & " Report" Then
    Me.vbxCrystal.Formulas(5) = "lblRptTitle = '" & lStr("Associations") & " Report'"
End If
Cri_Sorts = z% ' next section number to format

End Function

Private Function Cri_SortsDol()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_SortsDol = 0
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
        Case 1: grpField$ = "{tblDolEnt.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select

    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        If comGroup(1).ListIndex <> 0 Then
            Me.vbxCrystal.SectionFormat(0) = "GF4;F;X;X;X;X;X;X"
            Me.vbxCrystal.SectionFormat(1) = "GF2;F;X;X;X;X;X;X"
        End If
    Else
        grpCond$ = "GROUP" & CStr(2) & ";" & "{@EFullName}" & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    End If
    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblDolEnt.TB_DESC}"
        Case 1: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    Else
        
        grpCond$ = "GROUP" & CStr(3) & ";" & "{HRDOLENT.DE_TDATE}" & ";ANYCHANGE;D"
        If chkTerm.Value = True Then
            grpCond$ = "GROUP" & CStr(3) & ";" & "{HR_DOLENTWRK.ED_TDATE}" & ";ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblDolEnt.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    Else
        grpCond$ = "GROUP" & CStr(1) & ";" & "{@EFullName}" & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    End If
    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblDolEnt.TB_DESC}"
        Case 1: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    Else
        grpCond$ = "GROUP" & CStr(2) & ";" & "{HRDOLENT.DE_TDATE}" & ";ANYCHANGE;D"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    End If
    Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;F;X;X;X;X"
    Me.vbxCrystal.SectionFormat(3) = "GF3;F;X;F;X;X;X;X"
    Me.vbxCrystal.SectionFormat(4) = "GF4;F;X;F;X;X;X;X"
End If

Cri_SortsDol = z% ' next section number to format


End Function

Private Function Cri_SortsEduc()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
' imbeded in report
Cri_SortsEduc = 0
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0) <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    strSFormat$ = "GH1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = strSFormat$
    
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = strSFormat$
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREDU.EU_DEGREE}"
        Case 2: grpField$ = "(none)"
    End Select
    
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREDU.EU_DEGREE}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(3) & ";" & "{HREDU.EU_YEAR}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    End If
Else

    grpCond$ = "GROUP" & CStr(1) & ";" & "{HREDU.EU_COMPNO}" & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;F;X;X;X;X;X"

    'Here sorting starts when first grouping is "none"
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREDU.EU_DEGREE}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$ '0->1
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREDU.EU_DEGREE}"
            Case 1: grpField$ = "(none)"
        End Select

        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
            grpCond$ = "GROUP" & CStr(4) & ";" & "{HREDU.EU_YEAR}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(3) = grpCond$
        End If
    Else
        grpCond$ = "GROUP" & CStr(2) & ";" & "{HREDU.EU_COMPNO}" & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        'Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
        'Me.vbxCrystal.SectionFormat(1) = "GF1;F;F;X;X;X;X;X"
    
        'Here sorting starts when second grouping is also "none"
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{HREDU.EU_DEGREE}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
            grpCond$ = "GROUP" & CStr(4) & ";" & "{HREDU.EU_YEAR}" & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(3) = grpCond$
            'If third grouping is "none" then final sort is by "Education Year"
        Else
            grpCond$ = "GROUP" & CStr(3) & ";" & "{HREDU.EU_COMPNO}" & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
            grpCond$ = "GROUP" & CStr(4) & ";" & "{HREDU.EU_YEAR}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(3) = grpCond$
        End If
    End If
End If
   

Cri_SortsEduc = z% ' next section number to format


End Function

Private Function Cri_SortsOther()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_SortsOther = 0
grpField$ = getEGroup(comGroup(0).Text)
If grpField$ <> "(none)" Then
    
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    strSFormat$ = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;X;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1

    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblOEarn.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        Else
            GrpIdx% = comGroup(2).ListIndex
                Select Case GrpIdx%
                    Case 0: grpField$ = "{tblOEarn.TB_DESC}"
                    Case 1: grpField$ = "(none)"
                End Select
                If grpField$ <> "(none)" Then
                    grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
                    Me.vbxCrystal.GroupCondition(2) = grpCond$
                Else
                    grpCond$ = "GROUP" & CStr(3) & ";" & "{HREARN.TDATE}" & ";ANYCHANGE;D"
                    Me.vbxCrystal.GroupCondition(2) = grpCond$
                End If
        End If
        Else
    
        Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(1) = "GF1;F;F;X;X;X;X;X"
    'Else
    'Here sorting starts when first grouping is "none"
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblOEarn.TB_DESC}"
        Case 2: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblOEarn.TB_DESC}"
        Case 1: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        grpCond$ = "GROUP" & CStr(3) & ";" & "{HREARN.TDATE}" & ";ANYCHANGE;D"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    
     End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Else
        GrpIdx% = comGroup(1).ListIndex  'Second Selection
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{tblOEarn.TB_DESC}"
            Case 2: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
            
        GrpIdx% = comGroup(2).ListIndex  'Second Selection
        Select Case GrpIdx%
            Case 0: grpField$ = "{tblOEarn.TB_DESC}"
            Case 1: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & "{HREARN.TDATE}" & ";ANYCHANGE;D"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        End If
            
            Else
                GrpIdx% = comGroup(2).ListIndex  'Third Selection
                Select Case GrpIdx%
                Case 0: grpField$ = "{tblOEarn.TB_DESC}"
                Case 1: grpField$ = "(none)"
                End Select                       ' No Selection - Final Sort
                If grpField$ <> "(none)" Then
                dscGroup$ = comGroup(0).Text
                dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
                Me.vbxCrystal.Formulas(0) = dscGroup$
        
                grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond$
                Else
                    'dscGroup$ = comGroup(0).Text
                    'dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
                    'Me.vbxCrystal.Formulas(0) = dscGroup$
        
                    grpCond$ = "GROUP" & CStr(1) & ";" & "{HREARN.TDATE}" & ";ANYCHANGE;D"
                    Me.vbxCrystal.GroupCondition(0) = grpCond$
                End If
                Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
            End If
            Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;X;X;X;X;X"
            Me.vbxCrystal.SectionFormat(1) = "GF1;F;F;F;X;X;X;X"
            Me.vbxCrystal.SectionFormat(2) = "GF2;F;F;F;X;X;X;X"
            Me.vbxCrystal.SectionFormat(3) = "GF3;F;F;F;X;X;X;X"
            Me.vbxCrystal.SectionFormat(4) = "GF4;F;F;F;X;X;X;X"
        End If
       End If

Cri_SortsOther = z% ' next section number to format


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

For X% = 0 To 11
    If X <> 1 And X <> 2 And X <> 6 And X <> 7 And X <> 8 Then
        If Not clpCode(X).ListChecker Then Exit Function
    End If
Next X%

'If Len(txtCompPd) > 0 Then
'    If txtCompPd <> "Y" And txtCompPd <> "N" Then
'        MsgBox "Company Paid must be Y/N or blank"
'        txtCompPd.SetFocus
'        Exit Function
'    End If
'End If


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

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpDateRange(0)) > 0 And Len(dlpDateRange(1)) > 0 Then 'Hemu - 05/14/2003 Begin - Date should not be blank
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If 'Hemu - 05/14/2003 End - Date should not be blank

If Not clpCode(1).CheckList Then Exit Function
If Not clpCode(2).CheckList Then Exit Function

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
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRASSOC"

Screen.MousePointer = HOURGLASS

'Hide Date of Hire date range
lblDOH.Visible = False
dlpDateRange(2).Visible = False
dlpDateRange(3).Visible = False

If FormAssoc% = True Then
    frmRAssoc.Caption = lStr("Associations") & " Report"       'Laura Oct 20, 1997
    lblBCode.Caption = lStr("Associations")              'laura
    clpCode(1).Tag = "00-" & lStr("Associations") & " Codes"
    
    'txtCompPd.Visible = True 'laura
    'txtCompPd.Enabled = True 'laura
    'Label1.Visible = True 'laura
    'Label1.Enabled = True 'laura
    lblCompanyPd.Visible = False
    cmbCompanyPd.Visible = True
    cmbCompanyPd.Top = clpCode(5).Top
    cmbCompanyPd.Left = clpCode(5).Left + 320
    cmbCompanyPd.Clear
    cmbCompanyPd.AddItem ""
    cmbCompanyPd.AddItem "Y"
    cmbCompanyPd.AddItem "N"
    
    
    'Hemu
    lblCostOfEmp.Visible = False
    cmdCostOfEmp.Visible = False
    'Hemu
  
ElseIf FormDoll% = True Then
    frmRAssoc.Caption = "Dollar Entitlement Report" 'Laura oct 20, 1997
    lblBCode.Caption = "Entitlement Type"             'laura
    clpCode(1).Tag = "00-Enter Entitlement Type"
    'txtCompPd.Visible = False 'laura
    'txtCompPd.Enabled = False 'laura
    Label1.Visible = False 'laura
    Label1.Enabled = False 'laura
    lblShift.Left = Label1.Left
    txtShift.Left = 2115
    lblShift.Top = Label1.Top
    txtShift.Top = clpCode(5).Top
    'Hemu
    'Ticket #13188 - Begin
    lblCostOfEmp.Visible = True
    cmdCostOfEmp.Visible = True
    cmdCostOfEmp.Clear
    cmdCostOfEmp.AddItem ""
    cmdCostOfEmp.AddItem "Y"
    cmdCostOfEmp.AddItem "N"
    'Ticket #13188 - End
    'Hemu
    'If glbCompSerial = "S/N - 2288W" Then
    chkTerm.Visible = True
    'End If
      
    'Ticket #13188, Frank 06/11/07
    chkDolEntComment.Visible = True
    
    'Ticket #28789 - Show Amount Details
    chkDolEntDtls.Visible = True

ElseIf FormEduc% = True Then
    frmRAssoc.Caption = "Formal Education Report" 'Laura oct 20, 1997
    lblBCode.Caption = "Major Study"             'laura
    'Hemu -Begin - New control for Major study selection
    'clpCode(1).Tag = "00-Enter Major Study Code"      'laura
    clpCode(12).Visible = True
    clpCode(1).Visible = False
    'Hemu - End
    'txtCompPd.Visible = False 'laura
    'txtCompPd.Enabled = False 'laura
    Label1.Caption = "Minor Study"
    lblRenewal.Caption = "Degree"         'laura
    'dlpDateRange(0).Visible = False
    'dlpDateRange(1).Visible = False
    dlpDateRange(0).Left = 6560
    dlpDateRange(1).Left = 8220
    lblJOB.Visible = True
    lblFormalDate.Visible = True
    clpJOB.Visible = True
    lblComp.Visible = True
    cmdComp.Visible = True
    lblSchool.Visible = True
    clpCode(6).Visible = True
    clpCode(4).Visible = True
    clpCode(4).Tag = "00-Enter Degree Completed Code"
    clpCode(5).Visible = True
    clpCode(5).Tag = "00-Enter Minor Study Code"
    
    cmdComp.Clear
    cmdComp.AddItem ""
    cmdComp.AddItem "Y"
    cmdComp.AddItem "N"
    
    lblCompanyPd.Visible = True
    lblCompanyPd.Top = lblCostOfEmp.Top
    lblCompanyPd.Left = lblCostOfEmp.Left
    cmbCompanyPd.Visible = True
    cmbCompanyPd.Top = cmdCostOfEmp.Top
    cmbCompanyPd.Left = cmdCostOfEmp.Left
    cmbCompanyPd.Clear
    cmbCompanyPd.AddItem ""
    cmbCompanyPd.AddItem "Y"
    cmbCompanyPd.AddItem "N"
    
    'Label1.Visible = True 'laura
    'Label1.Enabled = True 'laura
    'txtCompPd.Visible = True
    'txtCompPd.Enabled = True
    
    'Hemu
    lblCostOfEmp.Visible = False
    cmdCostOfEmp.Visible = False
    'Hemu
    
    'Ticket #24906 - Vitalaire Report changes - DOH Date range
    If glbCompSerial = "S/N - 2380W" Then
        'Show Date of Hire date range
        lblDOH.Visible = True
        lblDOH.Caption = lStr("Original Hire")
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
    End If
    
ElseIf FormOther% = True Then
    frmRAssoc.Caption = "Other Earnings Report"
    lblBCode.Caption = "Earnings Type"
    clpCode(1).Tag = "00-Enter Earnings Type - Code"
    'txtCompPd.Visible = False
    'txtCompPd.Enabled = False
    Label1.Visible = False
    Label1.Enabled = False
    lblShift.Left = Label1.Left
    txtShift.Left = 2115
    lblShift.Top = Label1.Top
    txtShift.Top = clpCode(5).Top
    'Hemu
    'Ticket #13188 - Begin
    lblCostOfEmp.Visible = True
    cmdCostOfEmp.Visible = True
    cmdCostOfEmp.Clear
    cmdCostOfEmp.AddItem ""
    cmdCostOfEmp.AddItem "Y"
    cmdCostOfEmp.AddItem "N"
    'Ticket #13188 - End
    'Hemu
    
    'Ticket #24410 - City of Sarnia - Position Code added
    lblJOB.Visible = True
    clpJOB.Visible = True
End If          'laura

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad


If FormAssoc% = True Then 'laura
    clpCode(1).TablName = "TDCD"
ElseIf FormDoll% = True Then    'Laura
    clpCode(1).TablName = "EDOL"
ElseIf FormEduc% = True Then
    clpCode(4).TablName = "EUDE"
    clpCode(5).TablName = "EUMJ"
ElseIf FormOther% = True Then
    clpCode(1).TablName = "EARN"
End If

If glbLinamar Then clpCode(9).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(9).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)
Call setRptCaption(Me)

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
On Error GoTo Eh
    Dim c As Long
    
    If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
        panWindow.Height = Me.ScaleHeight - 200
        panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
                
        If panWindow.Height >= 9600 Then   '+ 230 Then
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form Resize", "Dollar Entitlement Report", "Form Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub Cri_Job()
Dim EECri As String, OneSet%, X%
Dim strCa$

If Len(clpJOB.Text) < 1 Then Exit Sub

If FormOther% = True Then
    strCa$ = "HREARN.OE_JOB"
Else
    strCa$ = "HR_JOB_HISTORY.JH_JOB"
End If

If Len(clpJOB.Text) > 0 Then
    EECri = EECri & "({" & strCa$ & "} in ['" & Replace(clpJOB.Text, ",", "','") & "']) " ' AND ({HR_JOB_HISTORY.JH_CURRENT}) "
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

Private Sub txtCompPd_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
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



