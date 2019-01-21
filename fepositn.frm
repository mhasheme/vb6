VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEPOSITION 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Position History"
   ClientHeight    =   11145
   ClientLeft      =   240
   ClientTop       =   735
   ClientWidth     =   12675
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11145
   ScaleWidth      =   12675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Frank Test"
      Height          =   495
      Left            =   11760
      TabIndex        =   180
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5295
      LargeChange     =   315
      Left            =   11800
      Max             =   100
      SmallChange     =   315
      TabIndex        =   62
      Top             =   2400
      Width           =   300
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fepositn.frx":0000
      Height          =   1725
      Left            =   0
      OleObjectBlob   =   "fepositn.frx":0014
      TabIndex        =   0
      Top             =   510
      Width           =   12135
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   72
      Top             =   10605
      Width           =   12675
      _Version        =   65536
      _ExtentX        =   22357
      _ExtentY        =   952
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdReCompDAccrual 
         Appearance      =   0  'Flat
         Caption         =   "&Re-Create Daily Accrual"
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
         Left            =   7560
         TabIndex        =   181
         Tag             =   "Employee's Daily Accrual as of Start Date"
         Top             =   15
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.CommandButton cmdJobFiles 
         Appearance      =   0  'Flat
         Caption         =   "&Job Files..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   134
         Tag             =   "Job Files related to this Job"
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton cmdBackupPosition 
         Appearance      =   0  'Flat
         Caption         =   "Backup Positions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   129
         Tag             =   "Edit Uder Defined Labels"
         Top             =   0
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton cmdEditLable 
         Appearance      =   0  'Flat
         Caption         =   "Edit &User Defined Labels"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   65
         Tag             =   "Edit Uder Defined Labels"
         Top             =   0
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CommandButton cmdPerform 
         Appearance      =   0  'Flat
         Caption         =   "Perfor&mance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   10335
         TabIndex        =   66
         Tag             =   "Call Performance Form"
         Top             =   330
         Visible         =   0   'False
         Width           =   1250
      End
      Begin Crystal.CrystalReport vbxCrystal2 
         Left            =   11760
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   12675
      _Version        =   65536
      _ExtentX        =   22357
      _ExtentY        =   873
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
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8280
         TabIndex        =   133
         Top             =   153
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Index           =   7
         Left            =   240
         TabIndex        =   71
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3120
         TabIndex        =   70
         Top             =   130
         Width           =   1740
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1440
         TabIndex        =   69
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label lblEmplNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Left            =   4440
         TabIndex        =   68
         Top             =   6030
         Width           =   1005
      End
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Left            =   0
      TabIndex        =   73
      Top             =   2280
      Width           =   11835
      Begin VB.Frame fraReptEDate 
         Height          =   1400
         Left            =   11520
         TabIndex        =   174
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
         Begin INFOHR_Controls.DateLookup dlpRptDate 
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   6
            Tag             =   "40-Enter Effective Date"
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpRptDate 
            Height          =   285
            Index           =   2
            Left            =   1170
            TabIndex        =   8
            Tag             =   "40-Enter Effective Date"
            Top             =   360
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpRptDate 
            Height          =   285
            Index           =   3
            Left            =   1170
            TabIndex        =   10
            Tag             =   "40-Enter Effective Date"
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpRptDate 
            Height          =   285
            Index           =   4
            Left            =   1170
            TabIndex        =   12
            Tag             =   "40-Enter Effective Date"
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   178
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   177
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   176
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   175
            Top             =   45
            Width           =   1020
         End
      End
      Begin VB.Frame frmLinamar 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   83
         Top             =   5730
         Visible         =   0   'False
         Width           =   8655
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "JH_LABOUREDATE"
            Height          =   285
            Index           =   0
            Left            =   6720
            TabIndex        =   51
            Tag             =   "40-Enter Effective Date"
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Frame frmWFCDIV 
            Height          =   330
            Left            =   1680
            TabIndex        =   171
            Top             =   0
            Width           =   3735
            Begin VB.TextBox txtLabCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataField       =   "JH_LabourCD"
               Height          =   285
               Left            =   320
               MaxLength       =   10
               TabIndex        =   50
               Tag             =   "00-Bonus Reporting #"
               Top             =   0
               Width           =   990
            End
            Begin VB.Label lblLabCodeDesc 
               Caption         =   "Unassigned"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1440
               TabIndex        =   172
               Top             =   0
               Width           =   2415
            End
            Begin VB.Image imgILabCode 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   0
               Picture         =   "fepositn.frx":8270
               Top             =   0
               Width           =   240
            End
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   49
            Top             =   150
            Visible         =   0   'False
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDLB"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour Code"
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
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   85
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   5520
            TabIndex        =   84
            Top             =   180
            Width           =   1020
         End
      End
      Begin VB.Frame frmNYCH 
         Height          =   495
         Left            =   9120
         TabIndex        =   161
         Top             =   7080
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox txtUSRLABEL3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   162
            Tag             =   "00-Code "
            Top             =   240
            Visible         =   0   'False
            Width           =   810
         End
         Begin INFOHR_Controls.CodeLookup clpSalDist 
            Height          =   285
            Left            =   1830
            TabIndex        =   163
            Top             =   0
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   8
         End
         Begin VB.Label lblSalDist 
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Distribution"
            Height          =   195
            Left            =   60
            TabIndex        =   164
            Top             =   45
            Width           =   1515
         End
      End
      Begin VB.Frame frmVitalAireJobFamily 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   9840
         TabIndex        =   151
         Top             =   5880
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   154
            Tag             =   "00-Bonus Reporting #"
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   153
            Tag             =   "00-Bonus Reporting #"
            Top             =   330
            Width           =   975
         End
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   152
            Tag             =   "00-Bonus Reporting #"
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Group Jobs"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   160
            Top             =   680
            Width           =   810
         End
         Begin VB.Label lblDouDivDesc 
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   159
            Top             =   660
            Width           =   4095
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   1305
            Picture         =   "fepositn.frx":83BA
            Top             =   660
            Width           =   240
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Job Family"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   158
            Top             =   350
            Width           =   1065
         End
         Begin VB.Label lblDouDivDesc 
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   157
            Top             =   330
            Width           =   4095
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   1305
            Picture         =   "fepositn.frx":8504
            Top             =   330
            Width           =   240
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Family"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   156
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblDouDivDesc 
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   155
            Top             =   0
            Width           =   4095
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   1305
            Picture         =   "fepositn.frx":864E
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.ComboBox comShift 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "00-Shift"
         Top             =   3450
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame frEssexLib 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   80
         TabIndex        =   145
         Top             =   5440
         Visible         =   0   'False
         Width           =   3090
         Begin MSMask.MaskEdBox medEssex 
            DataField       =   "JH_AVG_WHRS"
            Height          =   285
            Index           =   0
            Left            =   2075
            TabIndex        =   25
            Tag             =   "10-Average Weekly Timetable Hours"
            Top             =   0
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEssex 
            DataField       =   "JH_WROTATION"
            Height          =   285
            Index           =   1
            Left            =   2075
            TabIndex        =   26
            Tag             =   "10- Number of Rotation Weeks"
            Top             =   330
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEssex 
            DataField       =   "JH_DROTATION"
            Height          =   285
            Index           =   2
            Left            =   2075
            TabIndex        =   27
            Tag             =   "10-Number of Days within the Rotation"
            Top             =   660
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblEssexRotDays 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "# of Days within Rotation"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   148
            Top             =   705
            Width           =   1785
         End
         Begin VB.Label lblEssexRotWeeks 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "# of Rotation Weeks"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   147
            Top             =   375
            Width           =   1485
         End
         Begin VB.Label lblEssexAvgWkHrs 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Avg. Wkly. Timetable Hours"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   146
            Top             =   45
            Width           =   1980
         End
      End
      Begin VB.Frame frmSamuelProfitSharing 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8880
         TabIndex        =   140
         Top             =   6960
         Visible         =   0   'False
         Width           =   2325
         Begin VB.CheckBox chkProSha 
            Height          =   195
            Left            =   2160
            TabIndex        =   141
            Tag             =   "40-Lead Hand - y/n"
            Top             =   30
            Width           =   255
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Eligible for Profit Sharing"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   142
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU4"
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   138
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   2110
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmJobEnd 
         Height          =   660
         Left            =   6180
         TabIndex        =   130
         Top             =   3360
         Width           =   5415
         Begin INFOHR_Controls.DateLookup dlpENDDATE 
            DataField       =   "JH_ENDDATE"
            Height          =   285
            Left            =   1170
            TabIndex        =   44
            Tag             =   "41-Enter Position Start Date"
            Top             =   5
            Width           =   2800
            _ExtentX        =   4948
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "JH_ENDREAS"
            Height          =   285
            Index           =   2
            Left            =   1170
            TabIndex        =   45
            Tag             =   "01-End Reason Code"
            Top             =   320
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDRC"
         End
         Begin VB.Label lblReason 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   132
            Top             =   365
            Width           =   735
         End
         Begin VB.Label lblEndDATE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   131
            Top             =   50
            Width           =   885
         End
      End
      Begin VB.ComboBox cboShift 
         Height          =   315
         Left            =   3120
         TabIndex        =   18
         Top             =   3450
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10740
         TabIndex        =   127
         Top             =   4440
         Visible         =   0   'False
         Width           =   855
      End
      Begin Threed.SSCheck chkActPosition 
         Height          =   255
         Left            =   7380
         TabIndex        =   126
         Top             =   30
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Acting Position"
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
      Begin Threed.SSFrame fraPosition 
         Height          =   615
         Left            =   7380
         TabIndex        =   95
         Top             =   345
         Visible         =   0   'False
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   1085
         _StockProps     =   14
         Caption         =   "New Position"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optSalary 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   285
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   " New Salary"
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
         Begin Threed.SSOption optSalary 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   32
            Top             =   285
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   " Same Salary"
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
         Begin INFOHR_Controls.DateLookup dlpCurSEDate 
            Height          =   285
            Left            =   2280
            TabIndex        =   136
            Tag             =   "41-Enter Position Start Date"
            Top             =   600
            Visible         =   0   'False
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1210
            Enabled         =   0   'False
         End
         Begin Threed.SSOption optSalary 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   33
            Top             =   285
            Visible         =   0   'False
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Reporting Authority Change Only"
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
         Begin VB.Label lblCurSDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Salary Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   137
            Top             =   630
            Visible         =   0   'False
            Width           =   2085
         End
      End
      Begin VB.TextBox txtComments2 
         Appearance      =   0  'Flat
         DataField       =   "JH_COMMENT2"
         Height          =   285
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "00-Position Notes"
         Top             =   5100
         Width           =   2895
      End
      Begin VB.Frame frmMulti 
         Height          =   3800
         Left            =   6060
         TabIndex        =   89
         Top             =   600
         Visible         =   0   'False
         Width           =   5640
         Begin VB.TextBox txtPayrollID 
            Appearance      =   0  'Flat
            DataField       =   "JH_PAYROLL_ID"
            Height          =   285
            Left            =   1620
            MaxLength       =   25
            TabIndex        =   46
            Tag             =   "00-Payroll ID"
            Top             =   3420
            Width           =   1680
         End
         Begin VB.CheckBox chkUseForBenefit 
            Caption         =   "For Benefit"
            Height          =   315
            Left            =   3360
            TabIndex        =   125
            Top             =   3405
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtEmpType 
            BackColor       =   &H80000004&
            DataField       =   "JH_LEADHAND"
            Height          =   285
            Left            =   3960
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   2080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox comEmpType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1610
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Tag             =   "10-Type of Employee "
            Top             =   2065
            Visible         =   0   'False
            Width           =   2800
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "JH_EMP"
            Height          =   285
            Index           =   4
            Left            =   1290
            TabIndex        =   37
            Tag             =   "00-Employment Status - Code"
            Top             =   1120
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDEM"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "JH_ORG"
            Height          =   285
            Index           =   0
            Left            =   1290
            TabIndex        =   38
            Tag             =   "00-Union - Code"
            Top             =   1440
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin INFOHR_Controls.CodeLookup clpDiv 
            DataField       =   "JH_DIV"
            Height          =   285
            Left            =   1290
            TabIndex        =   34
            Tag             =   "00-Specific Division Desired"
            Top             =   170
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   1
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            DataField       =   "JH_DEPTNO"
            Height          =   285
            Left            =   1290
            TabIndex        =   35
            Top             =   490
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   7
            LookupType      =   2
         End
         Begin INFOHR_Controls.CodeLookup clpGLNum 
            DataField       =   "JH_GLNO"
            Height          =   285
            Left            =   1290
            TabIndex        =   36
            Tag             =   "00-General Ledger - Code"
            Top             =   810
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "JH_SECTION"
            Height          =   285
            Index           =   5
            Left            =   1290
            TabIndex        =   41
            Tag             =   "00-Section - Code"
            Top             =   2080
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpPT 
            DataField       =   "JH_PT"
            DataSource      =   " "
            Height          =   285
            Left            =   1290
            TabIndex        =   39
            Tag             =   "00-Category Codes"
            Top             =   1760
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDPT"
         End
         Begin INFOHR_Controls.CodeLookup clpRegion 
            DataField       =   "JH_REGION"
            DataSource      =   " "
            Height          =   285
            Left            =   1290
            TabIndex        =   43
            Tag             =   "00-Category Codes"
            Top             =   2430
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
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
            TabIndex        =   179
            Top             =   2475
            Width           =   510
         End
         Begin VB.Label lblPayID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   120
            Top             =   3465
            Width           =   675
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
            TabIndex        =   119
            Top             =   1805
            Width           =   765
         End
         Begin VB.Label lblSection 
            AutoSize        =   -1  'True
            Caption         =   "Section"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   2125
            Width           =   540
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Emp. Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   115
            Top             =   1165
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G/L Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   114
            Top             =   855
            Width           =   870
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   113
            Top             =   535
            Width           =   1560
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Division"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   112
            Top             =   215
            Width           =   555
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
            TabIndex        =   90
            Top             =   1485
            Width           =   510
         End
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         DataField       =   "JH_GRID"
         Height          =   285
         Left            =   1830
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGD"
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         DataField       =   "JH_SDATE"
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Tag             =   "41-Enter Position Start Date"
         Top             =   780
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "JH_SHIFT"
         Height          =   285
         Left            =   2145
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "00-Code assigned to the shift"
         Top             =   3450
         Width           =   810
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "JH_LDATE"
         Height          =   285
         Index           =   0
         Left            =   9480
         MaxLength       =   25
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "Ldate"
         Top             =   6030
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "JH_LTIME"
         Height          =   285
         Index           =   1
         Left            =   10080
         MaxLength       =   25
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "LTime"
         Top             =   6030
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "JH_LUSER"
         Height          =   285
         Index           =   2
         Left            =   11520
         MaxLength       =   25
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   "LUser"
         Top             =   4110
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         DataField       =   "JH_COMMENT"
         Height          =   285
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   23
         Tag             =   "00-Position Comments"
         Top             =   4770
         Width           =   2895
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   56
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1110
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU2"
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   63
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU3"
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   64
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1770
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmLinamar 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   80
         Top             =   6240
         Visible         =   0   'False
         Width           =   8655
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "JH_USREDATE"
            Height          =   285
            Index           =   1
            Left            =   6720
            TabIndex        =   54
            Tag             =   "40-Enter Effective Date"
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.TextBox txtLabel 
            Appearance      =   0  'Flat
            DataField       =   "JH_USRLABEL"
            Height          =   285
            Index           =   1
            Left            =   180
            MaxLength       =   20
            TabIndex        =   52
            Tag             =   "Enter Lable"
            Top             =   150
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CheckBox chkUserDef 
            DataField       =   "JH_USRCHECK"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   53
            Top             =   150
            Width           =   495
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   82
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   81
            Top             =   210
            Width           =   3015
         End
      End
      Begin VB.Frame frmLinamar 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   77
         Top             =   6780
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkUserDef 
            DataField       =   "JH_USRCHECK2"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   57
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox txtLabel 
            Appearance      =   0  'Flat
            DataField       =   "JH_USRLABEL2"
            Height          =   285
            Index           =   2
            Left            =   180
            MaxLength       =   20
            TabIndex        =   55
            Tag             =   "Enter Lable"
            Top             =   150
            Visible         =   0   'False
            Width           =   2895
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "JH_USREDATE2"
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   58
            Tag             =   "40-Enter Effective Date"
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   79
            Top             =   210
            Width           =   3015
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   5520
            TabIndex        =   78
            Top             =   180
            Width           =   1020
         End
      End
      Begin VB.Frame frmLinamar 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   74
         Top             =   7410
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtLabel 
            Appearance      =   0  'Flat
            DataField       =   "JH_USRLABEL3"
            Height          =   285
            Index           =   3
            Left            =   180
            MaxLength       =   20
            TabIndex        =   59
            Tag             =   "Enter Lable"
            Top             =   150
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CheckBox chkUserDef 
            DataField       =   "JH_USRCHECK3"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   60
            Top             =   180
            Width           =   495
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "JH_USREDATE3"
            Height          =   285
            Index           =   3
            Left            =   6720
            TabIndex        =   61
            Tag             =   "40-Enter Effective Date"
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblEdate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   5520
            TabIndex        =   76
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   75
            Top             =   210
            Width           =   2955
         End
      End
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   12600
         Top             =   4110
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   12600
         Top             =   3750
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_DHRS"
         Height          =   285
         Index           =   0
         Left            =   2145
         TabIndex        =   13
         Tag             =   "10-Usual working hours per day"
         Top             =   2460
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_WHRS"
         Height          =   285
         Index           =   1
         Left            =   2145
         TabIndex        =   14
         Tag             =   "10- Number of hours in work week"
         Top             =   2790
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_PHRS"
         Height          =   285
         Index           =   2
         Left            =   2145
         TabIndex        =   15
         Tag             =   "10-Usual working hours per pay period"
         Top             =   3120
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medFTENum 
         DataField       =   "JH_FTENUM"
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Tag             =   "10-Full - time equivalency"
         Top             =   4110
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medFTEHCalc 
         Height          =   285
         Left            =   2640
         TabIndex        =   94
         Top             =   4110
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSCheck chkCurrent 
         DataField       =   "JH_CURRENT"
         Height          =   285
         Index           =   0
         Left            =   9240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Current Position Record"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox medFTEHrs 
         DataField       =   "JH_FTEHRS"
         Height          =   285
         Left            =   2145
         TabIndex        =   22
         Tag             =   "10-FTE Hours worked per year"
         Top             =   4440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.00"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   12000
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JH_JREASON"
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   20
         Tag             =   "01-Reason for change in position - Code"
         Top             =   3780
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   2
         Left            =   1830
         TabIndex        =   9
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1770
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   7
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   5
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1110
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         DataField       =   "JH_JOB"
         Height          =   285
         Left            =   1830
         TabIndex        =   1
         Tag             =   "01-Position code"
         Top             =   120
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpPayrollCategory 
         DataField       =   "JH_PAYROLL_CATEGORY"
         Height          =   285
         Left            =   7560
         TabIndex        =   47
         Top             =   4770
         Visible         =   0   'False
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   9
      End
      Begin VB.Frame frmLinamar 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   86
         Top             =   5460
         Visible         =   0   'False
         Width           =   3885
         Begin VB.CheckBox chkLeadHand 
            Height          =   195
            Left            =   2115
            TabIndex        =   30
            Tag             =   "40-Lead Hand - y/n"
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblLeadHand 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "JH_LeadHand"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Hand"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   30
            TabIndex        =   87
            Top             =   0
            Width           =   795
         End
      End
      Begin Threed.SSCheck chkTrackCrsRenewal 
         DataField       =   "JH_TRK_CRS_RENEWAL"
         Height          =   285
         Left            =   7440
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   390
         Visible         =   0   'False
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Track Course Renewal"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   7320
         TabIndex        =   29
         Tag             =   "00-Band - Code"
         Top             =   3450
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin VB.Frame frmOCCAC 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -60
         TabIndex        =   117
         Top             =   5430
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox txtPosCtr 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "JH_POSITION_CONTROL"
            Height          =   285
            Left            =   2205
            MaxLength       =   6
            TabIndex        =   28
            Tag             =   "00-CCAC Position #"
            Top             =   30
            Width           =   1155
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   1950
            Picture         =   "fepositn.frx":8798
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lblPosCtr 
            Caption         =   "CCAC Position #"
            Height          =   345
            Left            =   120
            TabIndex        =   118
            Top             =   30
            Width           =   1575
         End
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   11
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   2115
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin MSMask.MaskEdBox medBillingRate 
         Height          =   285
         Left            =   7860
         TabIndex        =   143
         Tag             =   "10-Enter Billing Rate"
         Top             =   5430
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTotal 
         Height          =   285
         Left            =   9600
         TabIndex        =   149
         Tag             =   "21-Enter salary"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.DateLookup dlpDateSalE 
         Height          =   285
         Left            =   8880
         TabIndex        =   165
         TabStop         =   0   'False
         Tag             =   "40-Salary Effective Date"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   5370
         TabIndex        =   168
         Tag             =   "00-Employee Position Status  Code"
         Top             =   4440
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EPST"
      End
      Begin Threed.SSCheck chkPrimary 
         DataField       =   "JH_PRIMARY"
         Height          =   285
         Left            =   7380
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Primary Position"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   3120
         TabIndex        =   17
         Tag             =   "00-Shift"
         Top             =   3120
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SHFT"
         MaxLength       =   8
      End
      Begin VB.Image imgPosFilled 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   1480
         Picture         =   "fepositn.frx":88E2
         Stretch         =   -1  'True
         Top             =   1110
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgPosFilled 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   0
         Left            =   1480
         Picture         =   "fepositn.frx":8D24
         Stretch         =   -1  'True
         Top             =   450
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblWFCNote 
         Caption         =   $"fepositn.frx":9166
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5400
         TabIndex        =   173
         Top             =   1110
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label lblEStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Pos Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   169
         Top             =   4485
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblJobDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "JobDesc"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2145
         TabIndex        =   167
         Top             =   0
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblJob 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   166
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SalCode"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9600
         TabIndex        =   150
         Top             =   5160
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblBillingRate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Rate"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6090
         TabIndex        =   144
         Top             =   5475
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   139
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Image imgNoSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   10080
         Picture         =   "fepositn.frx":9210
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   10080
         Picture         =   "fepositn.frx":935A
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Caption         =   "Job Offer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   8970
         TabIndex        =   128
         Top             =   4440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblComment2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 2"
         Height          =   195
         Left            =   60
         TabIndex        =   124
         Top             =   5115
         Width           =   555
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Occupation"
         Height          =   195
         Left            =   6090
         TabIndex        =   123
         Top             =   5145
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label txtLambtonJob 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7860
         TabIndex        =   48
         Top             =   5100
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblPayrollCategory 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Payroll Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6090
         TabIndex        =   122
         Top             =   4815
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGrid 
         AutoSize        =   -1  'True
         Caption         =   "Grid Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   121
         Top             =   495
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
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
         Left            =   60
         TabIndex        =   111
         Top             =   165
         Width           =   1185
      End
      Begin VB.Label lblStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Left            =   60
         TabIndex        =   110
         Top             =   825
         Width           =   885
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   109
         Top             =   1155
         Width           =   1395
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   108
         Top             =   2505
         Width           =   780
      End
      Begin VB.Label lblHrsWeek 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Week"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   107
         Top             =   2835
         Width           =   930
      End
      Begin VB.Label lblHrsPayPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   106
         Top             =   3165
         Width           =   1260
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   105
         Top             =   3495
         Width           =   1725
      End
      Begin VB.Label lblEEStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Change"
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
         Left            =   60
         TabIndex        =   104
         Top             =   3825
         Width           =   1650
      End
      Begin VB.Label lblFTENum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE#"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   103
         Top             =   4155
         Width           =   480
      End
      Begin VB.Label lblFTEHrs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE Hours/Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   102
         Top             =   4485
         Width           =   1395
      End
      Begin VB.Label lblCompNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CompNo"
         DataField       =   "JH_COMPNO"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8400
         TabIndex        =   101
         Top             =   3180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 1"
         Height          =   195
         Left            =   60
         TabIndex        =   100
         Top             =   4785
         Width           =   555
      End
      Begin VB.Label lblBand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Band"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   99
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   98
         Top             =   1485
         Width           =   1290
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   97
         Top             =   1815
         Width           =   1290
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataField       =   "JH_EMPNBR"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9180
         TabIndex        =   96
         Top             =   3180
         Visible         =   0   'False
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmEPOSITION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim JobSnap_PayScale(20) As Double '15 -> 20 Ticket #24983 Franks 01/31/2014
Dim JobSnap_Salary_Code$
Dim JobSnap_MidPoint!
Dim fglbNew%
Dim savWHRS, savGrid, SavFte, SavFteHr, SavRpta(4), savSDate, savJOB As String
Dim savCurrent As Boolean
Dim Action

Dim fgtxtjob As String, fgtxtStartDate  As Variant
Dim fgtxtDhrs

Dim oPHRS, oWHRS, ODHRS, oJob As String, OSDATE
Dim OLeadHand, OLabourCD, OReason
Dim oPayrollID, oOrg, oDeptNo, oStatus, oGLNo
Dim oPayCategory

Dim OLambtonJob
Dim OFTE, fOldFTE, fNewFTE, fFTEDate
Dim oLABOUREDATE
Dim oENDDATE, oEndReason

Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim rsDATA2 As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim flgloaded As Boolean
Dim empPayrollID
Dim OBillingRate
Dim oSHIFT As String, oREPTAU As String
Dim locOrg As String
Dim oTrkCrsRen, flgNewCancel, flgTrainLstReset As Boolean
Dim locFTPT
Dim MailBody, MailBodyN, MailBodyR
Dim flgRehire As Boolean
Dim oPrimary As Boolean
Dim AbortTerm As Boolean
Dim xRept1FromDivMaster
Dim xIsWFCRetpEmpShowUp As Boolean
Dim IsWFC_CONP As Boolean
'frmMulti screen position: Left 4980 top 840

Private Function AUDITPSTN(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xPT, xDIV
Dim HRChangs As New Collection
Dim HRSalary As New Collection
Dim UpdateAudit As Boolean
Dim UptPositionDate As Date
Dim HRChangs1 As New Collection
Dim DCurSDate

'''On Error GoTo AUDIT_ERR

AUDITPSTN = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_DIV")) Then xDIV = "" Else xDIV = rsTB("ED_DIV")
Else
    xPT = ""
    xDIV = ""
End If

'Removed * Ticket #9899
Dim strFields As String
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_PHRS, AU_OLDPHRS, AU_WHRS, AU_OLDWHRS, AU_DHRS, AU_OLDDHRS, "
strFields = strFields & "AU_JOB, AU_SJDATE, AU_JREASON, AU_LEADHAND, AU_LABOURCD, AU_LABOUREDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_ORG, AU_BILLINGRATE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False
'Town of Aurora
If glbCompSerial = "S/N - 2378W" Then 'Or glbCompSerial = "S/N - 2375W" Then 'Or glbCompSerial = "S/N - 2363W" Then
    'Town of Aurora
    If glbCompSerial = "S/N - 2378W" Then   'Or glbCompSerial = "S/N - 2375W" Then
        'If Not IsNumeric(ODHRS) Then ODHRS = 0  'Ticket #20931 - as per mapping documentation - Transferred from Salary screen
        'If Not IsNumeric(medHours(0)) Then medHours(0) = 0  ''Ticket #20931 - as per mapping documentation - Transferred from Salary screen
        
        'If isChanged_Field(HRChangs, ODHRS, medHours(0), True) Then UpdateAudit = True ''Ticket #20931 - as per mapping documentation - Transferred from Salary screen
        If isChanged_Field(HRChangs, oPayCategory, clpPayrollCategory, False) Then UpdateAudit = True
        
        If chkCurrent(0) Or fglbNew Then
            Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
        End If
    End If

    UpdateAudit = False
Else
    If glbVadim And glbMulti Then
        If IsDate(glbChgTermDate) Then
            Call TermPayrollEmp(Date, glbLEE_ID, txtPayrollID.Text, Position)
        ElseIf (oJob = "" And chkCurrent(0)) And fglbNew Then
            If ifExistVadimPayrollID Then
                Call ReHireVadimEmp(Date, glbLEE_ID, txtPayrollID.Text)
            Else
                Call AddNewPayrollEmp(Position, dlpStartDate, glbLEE_ID, txtPayrollID.Text)
            End If
        End If
    End If
    UpdateAudit = False

    If Not IsNumeric(ODHRS) Then ODHRS = 0
    If Not IsNumeric(medHours(0)) Then medHours(0) = 0
    If Not IsNumeric(oWHRS) Then oWHRS = 0
    If Not IsNumeric(medHours(1)) Then medHours(1) = 0
    If Not IsNumeric(oPHRS) Then oPHRS = 0
    If Not IsNumeric(medHours(2)) Then medHours(2) = 0
    
    'Dim HRChangs As New Collection
    If isChanged_Field(HRChangs, ODHRS, medHours(0), True) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oWHRS, medHours(1), True) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oPHRS, medHours(2), True) Then UpdateAudit = True
    
    'Ticket #24565- DMuskoka - When the Hours/Pay Period changes, retransfer Employee's Salary (Total)
    If glbCompSerial = "S/N - 2373W" And oPHRS <> medHours(2) Then  'DMuskoka  - Pass Total which includes Premium
        lblSalCode.DataField = "SH_SALCD"
        medTotal.DataField = "SH_TOTAL"
        lblSalCode = GetSHData(glbLEE_ID, "SH_SALCD", "")
        medTotal = GetSHData(glbLEE_ID, "SH_TOTAL", "")
        
        If isChanged_Salary(HRSalary, "", medTotal, True) Then UpdateAudit = True
        If isChanged_Salary(HRSalary, "", lblSalCode) Then UpdateAudit = True
        
        Call Passing_Salary_Vadim(HRSalary, Salary, Date, medHours(2), medHours(1), glbLEE_ID, txtPayrollID.Text) 'txtPayrollID.Text)
        
        lblSalCode.DataField = ""
        medTotal.DataField = ""
    End If
    
    If isChanged_Field(HRChangs, oPayCategory, clpPayrollCategory, False) Then UpdateAudit = True
    
    If glbLambton Then
        If isChanged_Field(HRChangs, OLambtonJob, txtLambtonJob) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChangs, oJob, clpJob) Then UpdateAudit = True
    End If
    
    If OSDATE <> "" Then
        If isChanged_Field(HRChangs, Str(OSDATE), dlpStartDate) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChangs, OSDATE, dlpStartDate) Then UpdateAudit = True
    End If
    
    If isChanged_Field(HRChangs, OLeadHand, lblLeadHand) Then UpdateAudit = True
    'If isChanged_Field(HRChangs, OLabourCD, clpCode(3)) Then UpdateAudit = True
    If isChanged_Field(HRChangs, OLabourCD, txtLabCode) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oLABOUREDATE, dlpDate(0)) Then UpdateAudit = True
    If isChanged_Field(HRChangs, OReason, clpCode(1)) Then UpdateAudit = True
    
    If glbMulti Then
        If isChanged_Field(HRChangs, oPayrollID, txtPayrollID) Then UpdateAudit = True
        
        If glbVadim Then
            If isChanged_Field(HRChangs, oENDDATE, dlpENDDATE) Then UpdateAudit = True
            If isChanged_Field(HRChangs, oDeptNo, clpDept) Then UpdateAudit = True
            If isChanged_Field(HRChangs, oOrg, clpCode(0)) Then UpdateAudit = True
            If isChanged_Field(HRChangs, oGLNo, clpGLNum) Then UpdateAudit = True
            If isChanged_Field(HRChangs, oStatus, clpCode(4)) Then UpdateAudit = True
        End If
    End If
    
    If Not IsDate(glbChgTermDate) Or ((oJob = "" And chkCurrent(0)) And fglbNew) Then
        If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls - Ticket #14285
            If fglbNew Or CVDate(Format(dlpStartDate, "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy")) Then
                UptPositionDate = dlpStartDate
            Else
                UptPositionDate = Date
            End If
            Call Passing_Changes(HRChangs, Position, "M", UptPositionDate, glbLEE_ID, txtPayrollID.Text)
        Else
            'Ticket #28991 - Do not transfer from Previous Position screen (Not glbSetPos)
            If (chkCurrent(0) Or fglbNew) And Not glbSetPos Then   'Ticket #15751 - To prevent from updating Vadim with non-current position information.
                If chkCurrent(0) Then
                    'Current Position - transfer
                    Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
                Else
                    'Check if it's an older Position record being entered from Current Position screen
                    If Not glbMulti Then
                        DCurSDate = CurSDate()
                        If DCurSDate > 0 Then    '0 if no current record out there
                            DCurSDate = CVDate(DCurSDate)
                            If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) >= 0 Then
                                'Older Position record, do not transfer.
                            Else
                                'New Position record with latest Start Date - transfer.
                                Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
                            End If
                        Else
                            'Looks like first Position record so must be current
                            Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
                        End If
                    Else
                        'Multi Position record must be new current - transfer
                        Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
                    End If
                End If
            End If
        End If
    End If
    
    'Ticket #24565- DMuskoka - Transfer the Salary Effective Date at this time too if Same Salary New record, as Last Increment and Probation Date
    If glbCompSerial = "S/N - 2373W" Then
        If (chkCurrent(0) Or fglbNew) And optSalary(0).Value Then
            dlpDateSalE.DataField = "SH_EDATE"
            dlpDateSalE = dlpStartDate
            If isChanged_Field(HRChangs1, "", dlpDateSalE) Then UpdateAudit = True
            dlpDateSalE.DataField = ""
            Call Passing_Changes(HRChangs1, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
        End If
    End If
    
    If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
        If isChanged_Field(HRChangs, OBillingRate, medBillingRate) Then UpdateAudit = True
        If isChanged_Field(HRChangs, oENDDATE, dlpENDDATE) Then UpdateAudit = True
    End If
    
    If isChanged_Field(HRChangs, oSHIFT, txtShift) Then UpdateAudit = True
    If oREPTAU <> txtReptAuthority(0).Text Then UpdateAudit = True
    
End If

If UpdateAudit Then GoTo MODUPD Else GoTo MODNOUPD

GoTo MODNOUPD
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV
If oPHRS <> medHours(2) Then
    rsTA("AU_PHRS") = medHours(2)
    rsTA("AU_OLDPHRS") = oPHRS
End If
If oWHRS <> medHours(1) Then
    rsTA("AU_WHRS") = medHours(1)
    rsTA("AU_OLDWHRS") = oWHRS
End If
If ODHRS <> medHours(0) Then
    If Not IsNumeric(medHours(0)) Then medHours(0) = 0
    rsTA("AU_DHRS") = medHours(0)
    rsTA("AU_OLDDHRS") = ODHRS
End If
If glbInsync Then
    rsTA("AU_JOB") = clpJob.Text
    rsTA("AU_SJDATE") = dlpStartDate.Text
    rsTA("AU_JREASON") = clpCode(1).Text
Else
    If oJob <> clpJob.Text Then rsTA("AU_JOB") = clpJob.Text
    If oSHIFT <> txtShift.Text Then rsTA("AU_JOB") = clpJob.Text
    If oREPTAU <> txtReptAuthority(0).Text Then rsTA("AU_JOB") = clpJob.Text 'Ticket #12051 , for ADP interface (VitalAire, WFC, ...)
    If OSDATE <> dlpStartDate.Text Then rsTA("AU_SJDATE") = dlpStartDate.Text 'Ticket #12051
    If OReason <> clpCode(1).Text Then rsTA("AU_JREASON") = clpCode(1).Text
End If

If OLeadHand <> chkLeadHand Then rsTA("AU_LEADHAND") = lblLeadHand
'If OLabourCD <> clpCode(3).Text Then rsTA("AU_LABOURCD") = clpCode(3).Text
If OLabourCD <> txtLabCode.Text Then rsTA("AU_LABOURCD") = txtLabCode.Text

If oLABOUREDATE <> dlpDate(0).Text Then
    If IsDate(dlpDate(0).Text) Then
        rsTA("AU_LABOUREDATE") = dlpDate(0).Text
    Else
        rsTA("AU_LABOUREDATE") = Null
    End If
End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID

'Ticket #17853 - begin
'rsTA("AU_LDATE") = Date
'Ticket #20843 Franks 08/23/2011 - if the position start date < TODAY, always use TODAY FOR AU_LDATE
'If ACTX = "A" Then
'    rsTA("AU_LDATE") = dlpStartDate.Text
'Else
'If glbCompSerial = "S/N - 2436W" And ACTX = "A" Then   'Family Day  - Ticket #24009 Franks 07/18/2013
'    rsTA("AU_LDATE") = dlpStartDate.Text
'Else
    If CVDate(dlpStartDate.Text) > CVDate(Date) Then
        rsTA("AU_LDATE") = CVDate(dlpStartDate.Text)
    Else
        rsTA("AU_LDATE") = Date
    End If
'End If
'End If
'Ticket #17853 - end

rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
If glbMulti Then
    rsTA("AU_PAYROLL_ID") = txtPayrollID
Else
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    If isChanged_Field(HRChangs, ODHRS, medHours(0), True) _
    Or isChanged_Field(HRChangs, OBillingRate, medBillingRate) _
    Or isChanged_Field(HRChangs, oOrg, clpCode(0)) _
    Or isChanged_Field(HRChangs, oJob, clpJob) _
    Or isChanged_Field(HRChangs, oENDDATE, dlpENDDATE) _
    Then
        rsTA("AU_JOB") = clpJob.Text '# 7644
        If Len(clpCode(0)) > 0 Then
            rsTA("AU_ORG") = clpCode(0).Text
        End If
        If OBillingRate <> medBillingRate Then
            rsTA("AU_BILLINGRATE") = medBillingRate
        End If
    End If
End If
rsTA.Update

MODNOUPD:
AUDITPSTN = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js
Resume Next
End Function

Private Function ChkTermPos()

Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset

ChkTermPos = False

On Error GoTo JHS_Err
glbChgTermReason = ""
glbChgTermDate = ""
If IsDate(dlpENDDATE.Text) And oENDDATE = "" Then
    SQLQ = "Select JH_JOB,JH_EMPNBR,JH_ID FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & " AND JH_CURRENT <>0 "
    SQLQ = SQLQ & " AND JH_PAYROLL_ID ='" & txtPayrollID & "'"
    SQLQ = SQLQ & " AND JH_ID <> " & Data1.Recordset!JH_ID
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsTemp.EOF Then
        glbChgTermReason = "TERM"
        glbChgTermDate = dlpENDDATE.Text
    End If
End If
ChkTermPos = True
Exit Function
JHS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ChkTermPos", "HR_JOB_HISTORY", "SELECT")
Call RollBack '26July99 js
End Function

Private Function ifExistVadimPayrollID()
Dim X
Dim xBNo
Dim SQLQ
Dim rsVP As New ADODB.Recordset
On Error GoTo default_this
If Not glbVadim Then GoTo default_this
If txtPayrollID <> "" Then
    SQLQ = "SELECT EMP_NUM FROM EMPLOYEE WHERE EMP_NUM ='" & txtPayrollID & "'"
    rsVP.Open SQLQ, gdbPayroll, adOpenForwardOnly
    If Not rsVP.EOF Then
        ifExistVadimPayrollID = True
    Else
        ifExistVadimPayrollID = False
    End If
    rsVP.Close
    Exit Function
End If
default_this:
    ifExistVadimPayrollID = False
End Function

Private Function chkVadimPayrollID()
    If isTransfer(Position) Then
        chkVadimPayrollID = False
        If glbVadim And glbMulti Then
            If chkCurrent(0) And dlpENDDATE = "" And clpCode(2) = "" Then
                If IsDate(oENDDATE) And oEndReason <> "" Then
                    If Not ifExistVadimPayrollID Then
                        MsgBox "Vadim system does not have the Employee associated with the Payroll ID: " & txtPayrollID & "." & vbNewLine & "Please create a new position to add an new Employee in Vadim System."
                        Exit Function
                    Else
                        MsgBox "The Termination for the Payroll ID (" & txtPayrollID & ") is removed from Vadim system"
                        glbChgTermReason = oEndReason
                        glbChgTermDate = oENDDATE
                        Call ReHireVadimEmp(Date, glbLEE_ID, txtPayrollID)
                    End If
                End If
            End If
        End If
    End If
    chkVadimPayrollID = True
End Function

Private Function chkPosition()
Dim dd As Integer, DgDef As Double, Msg$, DCurSDate, DPrvSDate  As Variant
Dim Response%, X%
Dim rsEmp As New ADODB.Recordset
Dim CaseyFlag As Boolean
Dim xActPosFlag As Boolean
Dim xTempStr

chkPosition = False

xIsWFCRetpEmpShowUp = False

If glbCompSerial = "S/N - 2379W" Then 'Town of LaSalle Ticket #14534
    If Len(txtShift) = 0 Then
        txtShift.Text = "NOSD"
    End If
End If

If Len(clpJob.Text) <= 0 Then
    MsgBox "Position Code is required"
    clpJob.SetFocus
    Exit Function
Else
    If clpJob.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
        clpJob.SetFocus
        Exit Function
    End If
End If

If glbWFC Then 'Ticket #28340 Franks 03/21/2016
    If IsWFC_CONP Then  'Ticket #30376 Franks 07/17/2017
        'independent contractor
    Else
        If Mid(clpJob.Text, 5, 3) = "IND" Then
            MsgBox "Can't assign the independent contractor position to a regular WFC employee"
            Exit Function
        End If
    End If
    If fglbNew Then
        If IsInactivePos(clpJob.Text) Then
            MsgBox "'" & clpJob.Text & "' is Inactive Position Code. Please contact Corporate Total Rewards to review this Position Requirement."
            Exit Function
        End If
        If IsMissingBudPos(clpJob.Text) Then
            MsgBox "Please contact the info:HR corporate administrator to have them create the Budgeted Position Master for '" & clpJob.Text & "' "
            Exit Function
        End If
        If (glbUNION = "NONE" Or glbUNION = "EXEC") Then 'Salary employee only
                If optSalary(2).Value Then 'Reporting Authority Change Only
                    If SavRpta(0) = txtReptAuthority(0).Text And SavRpta(1) = txtReptAuthority(1).Text And SavRpta(2) = txtReptAuthority(2).Text And SavRpta(3) = txtReptAuthority(3).Text Then
                        Msg$ = "No change on Reporting Authority."
                        MsgBox Msg$
                        Exit Function
                    End If
                End If
                'If Not SavRpta(0) = txtReptAuthority(0).Text Then
                If fglbNew Then
                    If Len(txtReptAuthority(0).Text) > 0 Then
                        'xTempStr = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
                        'If Not xTempStr = txtReptAuthority(0).Text Then
                        If IsRept1PosNotMatchPosMaster(txtReptAuthority(0).Text, clpJob.Text) Then
                            glbMsgCustomVal = 11
                            frmMsgDialog.Show 1
                            'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
                            If glbMsgCustomVal = 2 Then 'If <<Cancel>> is checked, undo the change.
                                'Call cmdCancel_Click
                                txtReptAuthority(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
                                Exit Function
                            End If
                        End If
                        dlpRptDate(1).Text = Date 'Ticket #30491 Franks 09/07/2017
                    End If
                End If
                'End If
            'End If
            
            If NewHireForms.count = 0 Then 'for change only
                xIsWFCRetpEmpShowUp = True 'Ticket #29438 Franks 11/08/2016 - Salaried, New record, not new hire
            End If
        End If
    End If
    If Not fglbNew Then 'Ticket #29183 Franks 09/12/2016
        If Not savJOB = clpJob.Text Then
            'Msg$ = "Are you correcting a data entry error? "
            'Msg$ = Msg$ & Chr(10) & Chr(10) & "If it is not a data correction, cancel this transaction and create a new position record for this employee which will create accurate employee position history for this employee."
            Msg$ = "Are you entering a new position for this teammate? To maintain accuracy of Position Master, please click 'No' and click on the new record icon in the toolbar."
            Msg$ = Msg$ & Chr(10) & Chr(10) & "If this is data correction for this current position, click 'Yes'."
            
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg, 36, "Confirm")
            If Not Response% = 6 Then 'No
                clpJob.Text = savJOB
                Exit Function
            End If
        End If
        'If Not SavRpta(0) = txtReptAuthority(0).Text Then 'Ticket #29220 Franks 09/19/2016
        '    If Left(lblJobDesc.Caption, 1) = "U" Then
        '        'if the JOB CODE begins with a "U", ignore the RA#1 check.
        '    Else
        '        Msg$ = "Are you entering a new position for this teammate? To maintain accuracy of Position Master, please click 'No' and click on the new record icon in the toolbar."
        '        Msg$ = Msg$ & Chr(10) & Chr(10) & "If this is data correction for this current position, click 'Yes'."
        '        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
        '        Response% = MsgBox(Msg, 36, "Confirm")
        '        If Not Response% = 6 Then 'No
        '            'clpJob.Text = savJOB
        '            txtReptAuthority(0).Text = SavRpta(0)
        '            Exit Function
        '        End If
        '    End If
        'End If
        
        'Ticket #29343 Franks 10/17/2016 ------ begin
        If (glbUNION = "NONE" Or glbUNION = "EXEC") Then 'Salary employee only - replace "JOB CODE begins with a "U""
            xTempStr = ""
            If Not SavRpta(3) = txtReptAuthority(3).Text Then xTempStr = lblReptAuth(3).Caption
            If Not SavRpta(2) = txtReptAuthority(2).Text Then xTempStr = lblReptAuth(2).Caption
            If Not SavRpta(1) = txtReptAuthority(1).Text Then xTempStr = lblReptAuth(1).Caption
            If Not SavRpta(0) = txtReptAuthority(0).Text Then xTempStr = lblReptAuth(0).Caption
            If Len(xTempStr) > 0 Then
                Msg$ = "To change or delete " & xTempStr & ", a new record needs to be created. Click on the new record icon and select the Reporting Authority Change Only option."
                MsgBox Msg$
                If Not SavRpta(3) = txtReptAuthority(3).Text Then txtReptAuthority(3).Text = SavRpta(3)
                If Not SavRpta(2) = txtReptAuthority(2).Text Then txtReptAuthority(2).Text = SavRpta(2)
                If Not SavRpta(1) = txtReptAuthority(1).Text Then txtReptAuthority(1).Text = SavRpta(1)
                If Not SavRpta(0) = txtReptAuthority(0).Text Then txtReptAuthority(0).Text = SavRpta(0)
                Exit Function
            End If
        End If
        
        'Ticket #29343 Franks 10/17/2016 ------ end
    End If

End If

If glbMultiGrid Then
    If Len(clpGrid.Text) <= 0 Then
        MsgBox lStr("Grid Category is required")
        clpGrid.SetFocus
        Exit Function
    Else
        If clpGrid.Caption = "Unassigned" Then
            MsgBox lStr("Grid Category is required")
            clpGrid.SetFocus
            Exit Function
        End If
    End If
End If
If glbAdv Then
    If Not glbCompSerial = "S/N - 2242W" And Not glbCompSerial = "S/N - 2390W" And Not glbCompSerial = "S/N - 2418W" Then 'london ccac 'Collectcorp
        If isATIncluded(glbLEE_ID) Then
            If Len(txtShift.Text) = 0 Then
                MsgBox lStr("Shift is a required field")
                If txtShift.Visible Then
                    If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab
                        If Len(locOrg) > 0 Then txtShift.Text = locOrg ' txtShift = "NS" 'Ticket #14739
                    End If
                    txtShift.SetFocus
                End If
                Exit Function
            End If
        End If
    End If
End If
If glbVadim Then
    If glbMulti Then 'Ticket# 7751
        If Len(txtPayrollID.Text) = 0 Then
            MsgBox "Payroll ID is required"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    If Len(clpPayrollCategory.Text) < 1 Then
        MsgBox "Payroll Category is required field"
        clpPayrollCategory.SetFocus
        Exit Function
    Else
        If Not clpPayrollCategory.ListChecker Then
            clpPayrollCategory.SetFocus
            Exit Function
        End If
    End If
End If

If Len(dlpStartDate.Text) < 1 Then
    MsgBox "Position Start Date must be entered"
    dlpStartDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpStartDate.Text) Then
        MsgBox "Position Start Date is not a valid date"
        dlpStartDate.SetFocus
        Exit Function
    Else
        If glbSetPos Then
            DCurSDate = CurSDate()
            If DCurSDate > 0 Then    '0 if no current record out there
                DCurSDate = CVDate(DCurSDate)
                If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) <= 0 Then
                    Msg$ = "Warning...you cannot add or edit a record with a date"
                    'Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & " same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current position, "
                    'Msg$ = Msg$ & Chr(10) & "go to Position screen under Employee Menu."
                    Msg$ = Msg$ & "go to Position screen under Employee menu \ Work History/Compensation."
                    MsgBox Msg$, vbExclamation
                    dlpStartDate.SetFocus
                    Exit Function
                End If
            End If
        End If
        If Action = "A" Then
            If glbLinamar Then
                DCurSDate = CurSDate()
                If OReason = "NEWH" Then
                    If month(DCurSDate) = month(dlpStartDate.Text) And Year(DCurSDate) = Year(dlpStartDate.Text) And clpCode(1).Text <> "NEWH" Then
                        Msg$ = "Warning...you are creating a new position for this employee."
                        Msg$ = Msg$ & Chr(10) & "This employee was a New Hire this month."
                        Msg$ = Msg$ & Chr(10) & "To have the employee count as a New Hire in the HR Report, "
                        Msg$ = Msg$ & Chr(10) & "The Reason for Change on the record must be NEWH."
                        Msg$ = Msg$ & Chr(10) & "Click Yes to edit the Reason for Change."
                        Msg$ = Msg$ & Chr(10) & "Click No to update with the Reason for Change you entered."
                        If MsgBox(Msg$, vbYesNo) = vbYes Then
                            clpCode(1).SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        '-------
        rsEmp.Open "SELECT ED_EMPNBR,ED_DOB,ED_DOH,ED_DEPTNO,ED_PT,ED_BONUSDEPT,ED_SENDTE FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
        'Ticket #19965 - Samuel, Son & Co. Ltd. - Check against Seniority Date instead of Hire Date
        'Ticket #20910 - Friesens Corporation
        If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2279W" Then
            If CVDate(dlpStartDate.Text) < rsEmp("ED_SENDTE") Then
                If Not glbLambton Then
                    MsgBox lStr("Position Start Date must be greater than Seniority Date")
                    dlpStartDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            Else
                If CVDate(dlpStartDate.Text) <= rsEmp("ED_DOB") Then
                    MsgBox "Position Start Date must be greater than Date of Birth"
                    dlpStartDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            End If
        Else
            If CVDate(dlpStartDate.Text) < rsEmp("ED_DOH") Then
                If Not glbLambton Then
                    MsgBox lStr("Position Start Date must be greater than Original Hire Date")
                    dlpStartDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            Else
                If CVDate(dlpStartDate.Text) <= rsEmp("ED_DOB") Then
                    MsgBox "Position Start Date must be greater than Date of Birth"
                    dlpStartDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            End If
        End If
        
        If glbCompSerial = "S/N - 2214W" Then 'Casey House T#3359
            CaseyFlag = False
            If Not IsNull(rsEmp("ED_PT")) Then
                If rsEmp("ED_PT") = "FT" Then
                    CaseyFlag = True
                End If
            End If
            If Not IsNull(rsEmp("ED_DEPTNO")) Then
                If rsEmp("ED_DEPTNO") = "1425" Then
                    CaseyFlag = True
                End If
            End If
        End If
        
        If glbWFC Then
            If glbAdv And Not glbWFCFullRights Then 'Ticket #13867
                If IsNull(rsEmp("ED_BONUSDEPT")) Or Len(rsEmp("ED_BONUSDEPT")) = 0 Then
                    If Len(txtShift.Text) = 0 Then
                        MsgBox lStr("Shift is a required field")
                        txtShift.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
        rsEmp.Close
    End If
End If

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    If Len(medHours(2)) = 0 Then
        MsgBox "Hours/Per Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Reason Code is required"
    clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Reason Code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If glbLinamar Then
    'If Len(txtShift) < 1 Then
    '    MsgBox lStr("Shift is a required field")
    '    txtShift.SetFocus
    '    Exit Function
    'End If
    'Ticket #28846 Franks 07/14/2016 - replace txtShift with clpCode(8)
    If Len(clpCode(8).Text) < 1 Then
        MsgBox lStr("Shift is a required field")
        clpCode(8).SetFocus
        Exit Function
    End If
    'If Len(clpCode(3).Text) < 1 Then
    If Len(txtLabCode.Text) < 1 Then
        MsgBox "Labour Code is required field"
        'clpCode(3).SetFocus
        txtLabCode.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
    If Len(txtShift) < 1 Then
        MsgBox lStr("Shift") & " is a required field"
        If comShift.Visible And comShift.Enabled Then
            comShift.SetFocus
        End If
        Exit Function
    End If
End If

'----------Jaddy changed by Jerry request
If glbCompSerial = "S/N - 2347W" Then 'surrey place
    If Len(medFTENum) < 1 Then
        MsgBox "FTE # is required field"
        medFTENum.SetFocus
        Exit Function
    End If

End If
If glbCompSerial <> "S/N - 2276W" Then  'Temporarily removed for City of Niagara Falls - Jerry's request
    If glbVadim Then
        If Not glbLambton Then 'Ticket# 6692
            If Val(medHours(0)) = 0 Then
                MsgBox "Hours/Day is required field"
                medHours(0).SetFocus
                Exit Function
            End If
        End If
    End If
End If
If glbLambton Then
    If chkUseForBenefit And chkCurrent(0) Then
        If chkBenefitPayID Then
            MsgBox "Duplicate records found for ""For Benefit"". "
            chkUseForBenefit.SetFocus
            Exit Function
        End If
    End If
End If
If glbInsync Then
    If Not glbLambton And Not (glbCompSerial = "S/N - 2411W") Then   'Ticket# 6692
        '2411 Wellington-Dufferin-Guelph Public Health Ticket #16625
        If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
            If Val(medHours(0)) = 0 Then
                MsgBox "Hours/Day is required field"
                medHours(0).SetFocus
                Exit Function
            End If
            If Val(medHours(1)) = 0 Then
                MsgBox "Hours/Week is required field"
                medHours(1).SetFocus
                Exit Function
            End If
            If Val(medHours(2)) = 0 Then
                MsgBox "Hours/Pay Period is required field"
                medHours(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If

If glbCompSerial = "S/N - 2418W" Then 'charton hobbs
    If Not IsNumeric(medHours(0)) Then
        MsgBox "Hours/Day is required"
        medHours(0).SetFocus
        Exit Function
    End If
    If medHours(0) = 0 Then
        MsgBox "Hours/Day is required"
        medHours(0).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(1)) Then
        MsgBox "Hours/Week is required"
        medHours(1).SetFocus
        Exit Function
    End If
    If medHours(1) = 0 Then
        MsgBox "Hours/Week is required"
        medHours(1).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(2)) Then
        MsgBox "Hours/Per Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
    If medHours(2) = 0 Then
        MsgBox "Hours/Per Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
End If


'Ticket# 10389    'Burlington Tech & Linamar or Granite Club
If glbCompSerial = "S/N - 2351W" Or glbLinamar Or glbCompSerial = "S/N - 2241W" Then
    If Val(medHours(0)) = 0 Then
        MsgBox "Hours/Day is required field"
        medHours(0).SetFocus
        Exit Function
    End If
    If Val(medHours(1)) = 0 Then
        MsgBox "Hours/Week is required field"
        medHours(1).SetFocus
        Exit Function
    End If
    If Val(medHours(2)) = 0 Then
        MsgBox "Hours/Pay Period is required field"
        medHours(2).SetFocus
        Exit Function
    End If
End If
For X% = 0 To 3 '2
    If elpReptAuthShow(X%) = "0" Then elpReptAuthShow(X%) = ""
    If Len(elpReptAuthShow(X%)) > 0 Then
        If elpReptAuthShow(X%).Caption = "Unassigned" Then
            MsgBox "Rept. Authority Employee # not valid. Check Employee # and re-enter!"
            elpReptAuthShow(X%).SetFocus
            Exit Function
        End If
        If glbWFC Then 'Ticket #29343 Franks 10/17/2016
            If Len(dlpRptDate(X% + 1).Text) = 0 Then
                MsgBox "Effective Date is required if " & lblReptAuth(X%).Caption & " is entered."
                If dlpRptDate(X% + 1).Enabled Then dlpRptDate(X% + 1).SetFocus
                Exit Function
            End If
        End If
    End If
Next

'City of Pickering - Ticket #13281
If glbCompSerial = "S/N - 2217W" Or glbCompSerial = "S/N - 2386W" Then
    If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
        If Val(medHours(0)) = 0 Or medHours(0) = "" Then
            MsgBox "Hours/Day is required field"
            medHours(0).SetFocus
            Exit Function
        End If
        If Val(medHours(1)) = 0 Or medHours(1) = "" Then
            MsgBox "Hours/Week is required field"
            medHours(1).SetFocus
            Exit Function
        End If
        If Val(medHours(2)) = 0 Or medHours(2) = "" Then
            MsgBox "Hours/Pay Period is required field"
            medHours(2).SetFocus
            Exit Function
        End If
    End If
End If

'------------
DCurSDate = CurSDate()

If glbAddHisWarning And Action <> "M" And (Not glbSetPos) Then
    If DCurSDate > 0 Then    ' 0 if no current record out there
        DCurSDate = CVDate(DCurSDate)
        If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) >= 0 Then
            Msg$ = "Warning, you can not add a record with a date"
            Msg$ = Msg$ & Chr(10) & "the same or earlier than your most current record."
            'Msg$ = Msg & Chr(10) & "Do you want to proceed?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg$) ', DgDef, "Warning!")
            'If Response% = IDNO Then
            dlpStartDate.SetFocus
            Exit Function
            'End If
        End If
    End If
End If


If Len(medFTENum) > 0 Then
    If Not IsNumeric(medFTENum) Then
        medFTENum.SetFocus
        MsgBox "FTE# is invalid"
        Exit Function
    End If
End If

If Len(medFTEHrs) > 0 Then     'laura jan 05, 1997
    If Not IsNumeric(medFTEHrs) Then
        medFTEHrs.SetFocus
        MsgBox "FTE Hours/Year is invalid"
        Exit Function
    End If
End If

If glbMulti Then
    If glbWHSCC And glbLambton Then
        If Len(clpDiv) < 1 Then
            MsgBox lStr("Division is a required field")
            clpDiv.SetFocus
            Exit Function
        End If
        If Len(clpDept) < 1 Then
            MsgBox lStr("Department is a required field")
            clpDept.SetFocus
            Exit Function
        End If
        If Len(clpGLNum) < 1 Then
            MsgBox lStr("G/L # is a required field")
            clpGLNum.SetFocus
            Exit Function
        End If
        If Len(clpCode(4)) < 1 Then
            MsgBox lStr("Employment Status is a required field")
            clpCode(4).SetFocus
            Exit Function
        End If
        If Len(clpCode(0)) < 1 Then
            MsgBox lStr("Union Code is a required field")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    If clpDiv.Caption = "Unassigned" And Len(clpDiv.Text) > 0 Then
        MsgBox lStr("Division Code must be valid")
        clpDiv.SetFocus
        Exit Function
    End If

    If clpDept.Caption = "Unassigned" And Len(clpDept.Text) > 0 Then
        MsgBox lStr("Department Code must be valid")
         clpDept.SetFocus
        Exit Function
    End If
    
    If clpGLNum.Caption = "Unassigned" And Len(clpGLNum.Text) > 0 Then
        MsgBox lStr("G/L # Code must be valid")
         clpGLNum.SetFocus
        Exit Function
    End If
    
    ' danielk - 10/24/2002 - Added check to see if they actually put anything in the field
    If clpCode(4).Caption = "Unassigned" And Len(clpCode(4).Text) > 0 Then
        MsgBox lStr("Employment Status Code must be valid")
        clpCode(4).SetFocus
        Exit Function
    End If
    
    'wellington duffrine ticket ##17736
    If glbCompSerial = "S/N - 2411W" Then
        If Len(clpCode(0)) < 1 Then
            MsgBox lStr("Union Code is a required field")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    
    ' danielk - 10/24/2002 - Added check to see if they actually put anything in the field
    If clpCode(0).Caption = "Unassigned" And Len(clpCode(0).Text) > 0 Then
        MsgBox lStr("Union Code must be valid")
        clpCode(0).SetFocus
        Exit Function
    End If

    If Len(clpCode(5)) > 0 And clpCode(5).Caption = "Unassigned" Then
        MsgBox lStr("Section Code must be valid")
        clpCode(5).SetFocus
        Exit Function
    End If
    'Franks Aug 21,02 for Multi T#2743
    
    If Len(dlpENDDATE.Text) > 0 Then
        If IsDate(dlpENDDATE.Text) Then
            ' Except CITY OF PICKERING, Macaulay Child Dev. (Ticket #24564)
            If glbCompSerial <> "S/N - 2217W" And glbCompSerial <> "S/N - 2420W" Then
                chkCurrent(0) = False
            End If
            If DateDiff("d", dlpENDDATE.Text, dlpStartDate.Text) > 0 Then
                MsgBox "End Date must be later than Start Date"
                dlpENDDATE.SetFocus
                Exit Function
            End If
        Else
            MsgBox "End Date is invalid"
            dlpENDDATE.SetFocus
            Exit Function
        End If
    Else
        If Not glbOttawaCCAC Then
            If Not chkCurrent(0) Then
                MsgBox "End Date is required, if not current position"
                dlpENDDATE.SetFocus
                Exit Function
            End If
        End If
        
    End If
    If glbVadim And glbMulti Then
        If Not ChkTermPos Then Exit Function
    End If
End If

'Granite Club
If glbCompSerial = "S/N - 2241W" Then
    If Len(clpDiv) < 1 Then
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpDept) < 1 Then
        MsgBox lStr("Department is a required field")
        clpDept.SetFocus
        Exit Function
    End If
    If Len(clpCode(4)) < 1 Then
        MsgBox lStr("Employment Status is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(clpCode(0)) < 1 Then
        MsgBox lStr("Union Code is a required field")
        clpCode(0).SetFocus
        Exit Function
    End If
    If Len(clpPT) < 1 Then
        MsgBox lStr("Category Code is a required field")
        clpPT.SetFocus
        Exit Function
    End If
     If Len(clpRegion) < 1 Then
        MsgBox lStr("Region Code is a required field")
        clpRegion.SetFocus
        Exit Function
    End If
    'If Len(clpCode(5)) < 1 Then
    '    MsgBox lStr("Section Code is a required field")
    '    clpCode(5).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(2)) < 1 Then
    '    MsgBox lStr("Reason Code is a required field")
    '    clpCode(2).SetFocus
    '    Exit Function
    'End If
    
    If clpDiv.Caption = "Unassigned" And Len(clpDiv.Text) > 0 Then
        MsgBox lStr("Division Code must be valid")
        clpDiv.SetFocus
        Exit Function
    End If
    If clpDept.Caption = "Unassigned" And Len(clpDept.Text) > 0 Then
        MsgBox lStr("Department Code must be valid")
         clpDept.SetFocus
        Exit Function
    End If
    If clpCode(4).Caption = "Unassigned" And Len(clpCode(4).Text) > 0 Then
        MsgBox lStr("Employment Status Code must be valid")
        clpCode(4).SetFocus
        Exit Function
    End If
    
    If Len(clpCode(0)) < 1 Then
        MsgBox lStr("Union Code is a required field")
        clpCode(0).SetFocus
        Exit Function
    End If
    
    If clpPT.Caption = "Unassigned" And Len(clpPT.Text) > 0 Then
        MsgBox lStr("Category Code must be valid")
        clpPT.SetFocus
        Exit Function
    End If
    
    If clpRegion.Caption = "Unassigned" And Len(clpRegion.Text) > 0 Then
        MsgBox lStr("Region Code must be valid")
        clpRegion.SetFocus
        Exit Function
    End If
    
    'If Len(clpCode(5)) > 0 And clpCode(5).Caption = "Unassigned" Then
    '    MsgBox lStr("Section Code must be valid")
    '    clpCode(5).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(2)) > 0 And clpCode(2).Caption = "Unassigned" Then
    '    MsgBox lStr("Reason Code must be valid")
    '    clpCode(2).SetFocus
    '    Exit Function
    'End If
    'If Len(txtPayrollID.Text) = 0 Then
    '    MsgBox "Payroll ID is required"
    '    txtPayrollID.SetFocus
    '    Exit Function
    'End If
End If


If glbOttawaCCAC Then
    If GetSHData(glbLEE_ID, "SH_PAYP", "") = "E" Then
        If Not IsNumeric(medHours(1)) Then
            MsgBox "Hours/Week is required"
            medHours(1).SetFocus
            Exit Function
        End If
        If medHours(1) = 0 Then
            MsgBox "Hours/Week is required"
            medHours(1).SetFocus
            Exit Function
        End If
        If Not IsNumeric(medHours(2)) Then
            MsgBox "Hours/Per Period is required"
            medHours(2).SetFocus
            Exit Function
        End If
        If medHours(2) = 0 Then
            MsgBox "Hours/Per Period is required"
            medHours(2).SetFocus
            Exit Function
        End If
    End If
End If

'Franks Nov 6, 2002 for Casey House #3148
If glbCompSerial = "S/N - 2214W" Then
    If CaseyFlag Then
        If Not IsNumeric(medHours(0)) Then
            MsgBox "Hours/Day is required"
            medHours(0).SetFocus
            Exit Function
        End If
        If medHours(0) = 0 Then
            MsgBox "Hours/Day is required"
            medHours(0).SetFocus
            Exit Function
        End If
        If Not IsNumeric(medHours(2)) Then
            MsgBox "Hours/Per Period is required"
            medHours(2).SetFocus
            Exit Function
        End If
        If medHours(2) = 0 Then
            MsgBox "Hours/Per Period is required"
            medHours(2).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2375W" Then 'Timmins
    If lblHrsWeek.FontBold = True And Val(medHours(1)) = 0 Then
        MsgBox "Hours/Week is required field"
        medHours(1).SetFocus
        Exit Function
    End If
    
    If lblHrsPayPeriod.FontBold = True And Val(medHours(2)) = 0 Then
        MsgBox "Hours/Pay Period is required field"
        medHours(2).SetFocus
        Exit Function
    End If
End If

For X% = 0 To 2
    If Not IsNumeric(medHours(X%)) Then medHours(X%) = 0
Next

'Ticket #24565 - Making Hours/Pay Period mandatory as it's required to compute the Salary per Pay to transfer to
'Vadim as per the new formula
'Ticket #19113 - Making Hours/Week mandatory as it's required for computing Salary per Pay and transferring to Vadim
If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
    If Not IsNumeric(medHours(1)) Then
        MsgBox "Hours/Week is required"
        medHours(1).SetFocus
        Exit Function
    End If
    If medHours(1) = 0 Then
        MsgBox "Hours/Week is required"
        medHours(1).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(2)) Then
        MsgBox "Hours/Pay Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
    If medHours(2) = 0 Then
        MsgBox "Hours/Pay Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
End If

'7.9 - Enhancement - Open this City of Chatham-Kent logic for all
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    If (savSDate <> dlpStartDate.Text And Not fglbNew) Or fglbNew Then
        If Not fglbNew And Not chkCurrent(0) Then
            DCurSDate = CurSDate()
            If DCurSDate > 0 Then    '0 if no current record out there
                DCurSDate = CVDate(DCurSDate)
                'Ticket #24096 - Removing the check with same Start Date because Salary screen allows same
                'Effective Date Salary records
                'If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) <= 0 Then
                If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) < 0 Then
                    'MsgBox "Start Date cannot be same or later than your most recent current position record."
                    MsgBox "Start Date cannot be later than your most recent current position record."
                    dlpStartDate.SetFocus
                    Exit Function
                End If
            End If
        End If
    
        'Ticket #24096 - Removing the check with same Start Date because Salary screen allows same
        'Effective Date Salary records
        If Not glbSetPos And Not glbMulti Then    'If not Previous Position screen
            DPrvSDate = PrvSDate(IIf(chkCurrent(0) And fglbNew, False, chkCurrent(0).Value))
            If DPrvSDate > 0 Then    '0 if no current record out there
                'Ticket #24096
                'If DateDiff("d", DPrvSDate, CVDate(dlpStartDate.Text)) <= 0 Then
                If DateDiff("d", DPrvSDate, CVDate(dlpStartDate.Text)) = 0 Then
                    If glbWFC And optSalary(2).Value Then 'Ticket #29343 Franks 10/18/2016
                        'don't show this msg since "Reporting Authority Change Only" will create a new record with the same start date
                    Else
                        'MsgBox "Start Date cannot be same as or earlier than previous position(s)."
                        MsgBox "Start Date cannot be earlier than previous position(s)."
                        dlpStartDate.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    'What if reseting the Current flag to reset the Training List
    If Not fglbNew And savCurrent = True And savCurrent <> chkCurrent(0).Value And _
        Len(Trim(dlpENDDATE.Text)) = 0 And Len(Trim(clpCode(2).Text)) = 0 And _
        chkTrackCrsRenewal.Value = False And savSDate = dlpStartDate.Text Then
        
        flgTrainLstReset = True
    Else
        flgTrainLstReset = False
        
        '7.9 - Enhancement - For Friesens only for now as Chatham-Kent do not want this as well
        'When trying to save an existing record which is not current
        If chkCurrent(0).Value = False And (Not fglbNew) And (glbCompSerial = "S/N - 2279W") Then  'Not glbCompSerial = "S/N - 2188W" Then
            'Check first if End Date and End Reason had been entered for the older Position
            If Len(Trim(dlpENDDATE.Text)) = 0 Then
                MsgBox "End Date cannot be left blank for this previous Position"
                dlpENDDATE.SetFocus
                Exit Function
            ElseIf Not IsDate(dlpENDDATE.Text) Then
                MsgBox "Invalid End Date"
                dlpENDDATE.SetFocus
                Exit Function
            ElseIf CVDate(dlpENDDATE.Text) < CVDate(dlpStartDate.Text) Then
                MsgBox "End Date cannot be prior to Start Date"
                dlpENDDATE.SetFocus
                Exit Function
            ElseIf Len(Trim(clpCode(2).Text)) = 0 Then
                MsgBox "End Reason cannot be left blank for this previous Position"
                clpCode(2).SetFocus
                Exit Function
            ElseIf Not clpCode(2).ListChecker Then
                MsgBox "Invalid End Reason"
                clpCode(2).SetFocus
                Exit Function
            End If
        ElseIf chkCurrent(0) Then
            If IsDate(dlpENDDATE.Text) Then
                If CVDate(dlpENDDATE.Text) < CVDate(dlpStartDate.Text) Then
                    MsgBox "End Date cannot be prior to Start Date"
                    dlpENDDATE.SetFocus
                    Exit Function
                ElseIf (Len(Trim(clpCode(2).Text)) = 0 And glbWFC And GetEmpData(glbLEE_ID, "ED_EMP") = "CONP") Then
                    'Ticket #29660 - No Logic on End Date for Contract Employees with Current Position
                ElseIf (Len(Trim(clpCode(2).Text)) = 0 And glbCompSerial <> "S/N - 2420W") Then
                    'Macaulay Child Dev. (Ticket #24564) - End Date and No End Reason = Current
                    MsgBox "End Reason cannot be left blank for a Position with End Date"
                    clpCode(2).SetFocus
                    Exit Function
                ElseIf Len(Trim(clpCode(2).Text)) <> 0 And glbCompSerial = "S/N - 2420W" Then
                    'Macaulay Child Dev. (Ticket #24564) - End Date and End Reason = Not Current
                    chkCurrent(0) = False
                End If
            End If
        End If
    End If
'End If

If glbVadim Then
    If Not AUDITPSTN(Action) Then MsgBox "ERROR - AUDIT FILE"
Else
    If DCurSDate = 0 Then DCurSDate = dlpStartDate.Text  'New Record
    If IsDate(DCurSDate) Then  'Update Audit if Current Salary
        If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) <= 0 Then
            If Not AUDITPSTN(Action) Then MsgBox "ERROR - AUDIT FILE"
        End If
    End If
End If

If glbLinamar Then 'Ticket #28846 Franks 07/14/2016
    If Len(elpReptAuthShow(0).Text) = 0 Then
        MsgBox "Rept. Authority 1 is required."
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2382W" Then 'Ticket #18090 Samuel
    If Len(elpReptAuthShow(0).Text) = 0 Then
        MsgBox "Rept. Authority 1 is required."
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
    'Samuel Ticket #20371 Franks 05/25/2011
    If fglbNew Then
        If optSalary(0).Value Then 'Same Salary
            frmMsgDialog.Show 1
        End If
    End If
    
    'Ticket #20886 - begin Franks 08/31/2011
    If lblReptAuth(3).FontBold Then
        If Len(elpReptAuthShow(3).Text) = 0 Then
            MsgBox "Rept. Authority 4 is required."
            elpReptAuthShow(3).SetFocus
            Exit Function
        End If
    End If
    If lblTitle(1).FontBold Then
        If chkProSha.Value = 0 Then
            MsgBox "Eigible for Profit Sharing must be checked."
            chkProSha.SetFocus
            Exit Function
        End If
    End If
    'Ticket #20886 - end
    
    'Ticket #20885 Franks 11/11/2011 - begin
    If fglbNew Then
        If NewHireForms.count = 0 Then 'for change only
            If Not glbtermopen Then 'active only
                Call CheckReptAuth
            End If
        End If
    End If
    'Ticket #20885 Franks 11/11/2011 - end
End If

'Four Villages Community Health Centre - Ticket #18221
If glbCompSerial = "S/N - 2425W" Then
    If Not IsNumeric(medHours(2)) Or medHours(2) = 0 Then
        MsgBox "Hours/Per Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
End If

If glbWFC Then
    'Ticket #24767 Franks 12/10/2013 - begin
    xTempStr = isValidWFCJob(clpJob.Text, glbLEE_ID)
    If xTempStr = 1 Then
        'Ticket #24817 Franks 12/16/2013
        'MsgBox "If the employee's union code is not NONE or EXEC and their Category is FT, the Position Status must say 'REG' "
        MsgBox "Position selected in not valid for an hourly employee."
        clpJob.SetFocus
        Exit Function
    End If
    If xTempStr = 2 Then
        'Ticket #24817 Franks 12/16/2013
        'MsgBox "If the employee's union code is NONE or EXEC and their Category is FT, the Position Status must not say 'REG' "
        MsgBox "Position selected in not valid for a salaried employee."
        clpJob.SetFocus
        Exit Function
    End If
    'Ticket #24767 Franks 12/10/2013 - end
    
    If Len(txtShift) = 0 Then
        txtShift = "NS"
    End If
    
    If Len(elpReptAuthShow(0)) = 0 Then
        MsgBox "Rept. Authority 1 is required."
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(0)) Then
        MsgBox "Hours/Day is required"
        medHours(0).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(1)) Then
        MsgBox "Hours/Week is required"
        medHours(1).SetFocus
        Exit Function
    End If
    If Not IsNumeric(medHours(2)) Then
        MsgBox "Hours/Per Period is required"
        medHours(2).SetFocus
        Exit Function
    End If
    
    Call WFC_PT_PenChanged  'Ticket #22991 Franks 12/24/2012
    
End If

If glbCompSerial = "S/N - 2259W" Then 'Oxford Ticket #17400
    If glbMulti Then
        'Ticket #21599 Franks 03/01/2012
        If fglbNew Then 'new record
            If CheckDuplCurrent(glbLEE_ID, clpJob.Text) Then
                Msg$ = "There is another current position for the same Position Code '" & clpJob.Text & "' " & Chr(10)
                Msg$ = Msg$ & "You can't have two current positions for the same Code" & Chr(10)
                Msg$ = Msg$ & "Please uncheck the Current Position Record flag for the previous Current Position." & Chr(10)
                MsgBox Msg$
                Exit Function
            End If
        End If
    Else
        If chkCurrent(0).Value = True Then 'current position
            If chkActPosition.Value = False Then
                chkActPosition.Value = True
            End If
        End If
    End If
End If

'Ticket #20105 Franks 09/20/2011
If glbCompSerial = "S/N - 2384W" Then 'Town of St. Marys
    If fglbNew = False Then 'not new record
        If Not glbtermopen Then
            If chkActPosition.Value = False Then
                'check if there is a "Acting Position" checked
                xActPosFlag = checkActPosEmp(glbLEE_ID, Data1.Recordset("JH_ID"))
                If xActPosFlag Then
                    MsgBox "There is no " & chkActPosition.Caption & " checked. " & Chr(10) & " Please select one current position as " & chkActPosition.Caption
                End If
            End If
        End If
    End If
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti And chkPrimary Then
    'Primary Position can only be assigned to Current Position and only one Primary Position
    'Acting Position cannot be Primary Position
    If chkCurrent(0) And chkActPosition.Value = True Then
        MsgBox "An " & chkActPosition.Caption & " cannot be a Primary Position as well."
        chkPrimary.SetFocus
        Exit Function
    End If
    If Not chkCurrent(0) Then
        MsgBox "A Primary Position must be a Current Position as well."
        chkPrimary.SetFocus
        Exit Function
    End If
    
    'Check if another Current Position is already selected as Primary Position
    If fglbNew = False Then
        If PrimaryPositionExists(glbLEE_ID, Data1.Recordset("JH_ID")) Then
            MsgBox "Another Primary Position already exists. You cannot have more than one Primary Position.", vbExclamation
            chkPrimary.SetFocus
            Exit Function
        End If
    Else
        If PrimaryPositionExists(glbLEE_ID) Then
            MsgBox "Another Primary Position already exists. You cannot have more than one Primary Position.", vbExclamation
            chkPrimary.SetFocus
            Exit Function
        End If
    End If
End If

chkPosition = True

End Function

Private Function checkActPosEmp(xEmpNo, xJH_ID)
Dim rsLocEmpJob As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT JH_EMPNBR, JH_ID, JH_POSITION_CONTROL FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & " AND JH_CURRENT <>0 "
    SQLQ = SQLQ & " AND NOT JH_ID = " & xJH_ID & " "
    SQLQ = SQLQ & " AND JH_POSITION_CONTROL = 'YES' "
    rsLocEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsLocEmpJob.EOF Then
        retVal = True
    End If
    checkActPosEmp = retVal
End Function

Private Sub cboShift_Change()
    'If cboShift.ListIndex > -1 Then
        txtShift.Text = cboShift.Text
    'End If
End Sub

Private Sub cboShift_Click()
    'If cboShift.ListIndex > -1 Then
        txtShift.Text = cboShift.Text
    'End If
End Sub

Private Sub chkCurrent_Click(Index As Integer, Value As Integer)
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        If chkCurrent(0).Value = True Then
            chkTrackCrsRenewal.Visible = False
        Else
            chkTrackCrsRenewal.Visible = True
        End If
    End If
    
    If chkCurrent(0) Then
        chkActPosition.Enabled = True
        
        'WDGPHU - Ticket #27899
        If glbCompSerial = "S/N - 2411W" And glbMulti Then
            chkPrimary.Enabled = True
        End If
    Else
        chkActPosition.Enabled = False
        
        'WDGPHU - Ticket #27899
        If glbCompSerial = "S/N - 2411W" And glbMulti Then
            chkPrimary.Enabled = False
        End If
    End If
    
End Sub

Private Sub chkLeadHand_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkLeadHand_LostFocus()
If chkLeadHand.Value = 1 Then
    lblLeadHand.Caption = "Y"
Else
    lblLeadHand.Caption = "N"
End If
End Sub

Private Sub clpGrid_LostFocus()
Call Job_Desc
End Sub

Private Sub clpJob_Change()
Call setGridList
If glbWFC Then 'Ticket #27820 Franks 11/26/2015
    lblJobDesc.Caption = GetJobData(clpJob.Text, "JB_JOBCODE")
End If
End Sub

Private Sub clpJob_LostFocus()
Call Job_Desc

'Ticket #16212 - Remove this logic because on Position Master it contains Hour/Pay Period
'If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
'    'Get the Hours per Day from HRJOB
'    medHours(0).Text = Get_DayHours_for_Job(clpJob.Text)
'End If

If glbWFC Then 'Ticket #29069 Franks 08/18/2016
    Call WFCReptDisp
End If

End Sub

Private Sub WFCReptDisp()
If glbWFC Then 'Ticket #29069 Franks 08/18/2016
    xRept1FromDivMaster = ""
    If fglbNew Then
        If Len(txtReptAuthority(0).Text) = 0 And Len(clpJob.Text) > 0 Then
            txtReptAuthority(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
            'Ticket #29183 Franks 09/12/2016 - begin
            If Len(txtReptAuthority(0).Text) > 0 Then
                xRept1FromDivMaster = txtReptAuthority(0).Text
            End If
            If Len(txtReptAuthority(0).Text) = 0 Then
                lblWFCNote.Top = 4110 '1110
                lblWFCNote.Left = 5400
                lblWFCNote.Visible = True
            End If
            'Ticket #29183 Franks 09/12/2016 - end
        End If
        Call WFCReptDateSetup
    Else
        lblWFCNote.Visible = False
    End If
End If
End Sub

Private Sub cmdJobFiles_Click()
    glbPos = clpJob.Text
    glbDocName = "EmpPosJobFiles"
    frmJobDocument.Show 1
    DoEvents
End Sub

Private Sub cmdBackupPosition_Click()
    Unload Me
    Load frmEPositionBK
End Sub

Private Sub cmdEditLable_Click()
Dim X
For X = 1 To 3
    txtLabel(X).Visible = True
Next
If txtLabel(1) = "" Then txtLabel(1) = "POC"
If txtLabel(2) = "" Then txtLabel(2) = "LPS Program Manger"
If txtLabel(3) = "" Then txtLabel(3) = "P.I. Program"

End Sub

Private Sub cmdJobFiles_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdReCompDAccrual_Click()
    Dim Response%
    Dim xFromDate, xToDate
    
    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If cmdReCompDAccrual.Visible = True Then
        If glbCompEntVacDaily Then
            'Check if the Start Date is within the entitlement period. If not then cannot re-create the Daily Accrual
            'Get Entitlement Period of the employee
            xFromDate = GetEmpData(glbLEE_ID, "ED_EFDATE")
            xToDate = GetEmpData(glbLEE_ID, "ED_ETDATE")
            
            'If the Start Date is within the entitlement period then only Daily Accrual update will take place.
            If IsDate(xFromDate) And IsDate(xToDate) Then
                If CVDate(dlpStartDate.Text) >= CVDate(xFromDate) And CVDate(dlpStartDate.Text) <= CVDate(xToDate) Then
                    'Comfirm the Re-Computation of Daily Accrual
                    Response% = MsgBox("This function will create/recreate the Daily Accruals for this Employee as of the Position Start Date." & Chr(10) & Chr(10) & "Are you sure you want to proceed with this?", vbQuestion + vbYesNo, "Create the Daily Accrual File")
                    If Response% = IDNO Then
                        Exit Sub
                    End If
                    
                    'Re-Create Daily Accrual for this Employee
                    Call Recompute_DailyAccrualFile(glbLEE_ID, dlpStartDate.Text)
                    
                    MsgBox ("Daily Accrual created for this employee successfully."), vbInformation, "Daily Accrual Created"
                Else
                    MsgBox ("Daily Accrual cannot be re-created. This employee's Start Date is outside the Entitlement Period."), vbExclamation, "Failed to Create Daily Accrual"
                End If
            Else
                MsgBox ("Daily Accrual cannot be re-created. This employee's Entitlement Period is invalid."), vbExclamation, "Failed to Create Daily Accrual"
            End If
        End If
    End If
End Sub

Private Sub cmdReCompDAccrual_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Command1_Click() 'for testing only
'If glbWFC Then 'Ticket #29438 Franks 11/08/2016
    If IsWFCReptAuth(glbLEE_ID, "") Then
        glbWFC_IPPopFormName = "WFCEmpListWithRept"
        glbWFC_IncePlanID = glbLEE_ID 'Employee based
        'glbWFC_IncePlanID = -100 'Position Master based
        frmCheckListView.lblStDate = dlpStartDate.Text
        frmCheckListView.Show 1
    End If
'End If
End Sub

Private Sub comShift_Click()
    If glbCompSerial = "S/N - 2380W" Then 'Vitalaire
        'Ticket #24976 - Label changed, and add dropdown list
        If comShift.ListIndex <> -1 Then
            txtShift.Text = Trim(Left(comShift.Text, InStr(1, comShift.Text, "-") - 1))
        End If
    End If
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
        If comShift.ListIndex <> -1 Then
            txtShift.Text = Trim(Left(comShift.Text, InStr(1, comShift.Text, "-") - 1))
        End If
    End If
End Sub

Private Sub dlpStartDate_LostFocus()
Call WFCReptDateSetup 'Ticket #29343 Franks 10/17/2016
End Sub

Private Sub elpReptAuthShow_LostFocus(Index As Integer)
    Call WFCReptDateSetup 'Ticket #29343 Franks 10/17/2016
End Sub

Private Sub WFCReptDateSetup() 'Ticket #29343 Franks 10/17/2016
If glbWFC Then
    If fglbNew Then 'new record
        If IsDate(dlpStartDate.Text) Then
            If Len(dlpRptDate(1).Text) = 0 Then
                If Len(elpReptAuthShow(0).Text) > 0 Then dlpRptDate(1).Text = dlpStartDate.Text Else dlpRptDate(1).Text = ""
            End If
            If Len(dlpRptDate(2).Text) = 0 Then
                If Len(elpReptAuthShow(1).Text) > 0 Then dlpRptDate(2).Text = dlpStartDate.Text Else dlpRptDate(2).Text = ""
            End If
            If Len(dlpRptDate(3).Text) = 0 Then
                If Len(elpReptAuthShow(2).Text) > 0 Then dlpRptDate(3).Text = dlpStartDate.Text Else dlpRptDate(3).Text = ""
            End If
            If Len(dlpRptDate(4).Text) = 0 Then
                If Len(elpReptAuthShow(3).Text) > 0 Then dlpRptDate(4).Text = dlpStartDate.Text Else dlpRptDate(4).Text = ""
            End If
        End If
        'Ticket #30491 Franks 09/07/2017 - begin
        If Len(SavRpta(0)) > 0 And Len(elpReptAuthShow(0).Text) > 0 Then 'change rept
            If Not SavRpta(0) = txtReptAuthority(0).Text Then
                dlpRptDate(1).Text = Date
            End If
        End If
        If Len(SavRpta(1)) > 0 And Len(elpReptAuthShow(1).Text) > 0 Then 'change rept
            If Not SavRpta(1) = txtReptAuthority(1).Text Then
                dlpRptDate(2).Text = Date
            End If
        End If
        If Len(SavRpta(2)) > 0 And Len(elpReptAuthShow(2).Text) > 0 Then 'change rept
            If Not SavRpta(2) = txtReptAuthority(2).Text Then
                dlpRptDate(3).Text = Date
            End If
        End If
        If Len(SavRpta(3)) > 0 And Len(elpReptAuthShow(3).Text) > 0 Then 'change rept
            If Not SavRpta(3) = txtReptAuthority(3).Text Then
                dlpRptDate(4).Text = Date
            End If
        End If
        'Ticket #30491 Franks 09/07/2017 - end
    End If
    If Not fglbNew Then 'modified
        'If IsDate(dlpStartDate.Text) Then
            If Len(SavRpta(0)) = 0 And Len(elpReptAuthShow(0).Text) > 0 Then 'add a rept
                If Len(dlpRptDate(1).Text) = 0 Then dlpRptDate(1).Text = Date
            End If
            If Len(SavRpta(1)) = 0 And Len(elpReptAuthShow(1).Text) > 0 Then 'add a rept
                If Len(dlpRptDate(2).Text) = 0 Then dlpRptDate(2).Text = Date
            End If
            If Len(SavRpta(2)) = 0 And Len(elpReptAuthShow(2).Text) > 0 Then 'add a rept
                If Len(dlpRptDate(3).Text) = 0 Then dlpRptDate(3).Text = Date
            End If
            If Len(SavRpta(3)) = 0 And Len(elpReptAuthShow(3).Text) > 0 Then 'add a rept
                If Len(dlpRptDate(4).Text) = 0 Then dlpRptDate(4).Text = Date
            End If
        'End If
    End If
End If
End Sub

Private Sub imgILabCode_Click()
Call txtLabCode_DblClick
End Sub

Private Sub imgPosFilled_Click(Index As Integer)
Dim xMsg As String
    If Index = 0 Then
        Call getReptPosEmpListByPos(clpJob.Text, "")
    End If
    If Index = 1 Then
        Call getReptPosEmpListByEmp(elpReptAuthShow(0).Text)
    End If
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEPOSITION")
    Call FillMemoFile(SQLQ, "Offer")
End Sub

Private Sub cmdImport_Click()
    If fglbNew Then
        If Len(Trim(clpJob.Text)) = 0 Or Not IsDate(dlpStartDate.Text) Then
            MsgBox "Position Code and Start Date is mandatory to attach a document", vbInformation, "Invalid Position Code / Start Date"
            Exit Sub
        End If
        glbJob = clpJob.Text
        glbSDate = dlpStartDate.Text
    End If
    glbDocNewRecord = fglbNew
    glbDocName = "Offer"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEPOSITION")
End Sub

Private Sub elpReptAuthShow_Change(Index As Integer)
txtReptAuthority(Index).Text = getEmpnbr(elpReptAuthShow(Index).Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Keepfocus As Boolean
    If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    Keepfocus = Not isUpdated(Me)
    Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
    fraDetail.Height = 6000 '5535 '5000
    If glbLinamar Then fraDetail.Height = 7900 '7815
    If glbOttawaCCAC Then fraDetail.Height = 6000
    If glbCompSerial = "S/N - 2296W" Then fraDetail.Height = 7000   'Essex County Library
    
    If Me.Height >= vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + panControls.Height + 550 Then
        scrControl.Value = 0
        fraDetail.Top = vbxTrueGrid.Height + panEEDESC.Height + 60
        scrControl.Visible = False
        Exit Sub
    End If
    If Me.Height < vbxTrueGrid.Height + panEEDESC.Height + scrControl.Top + panControls.Height + 400 Then Exit Sub
    scrControl.Visible = True
    
    scrControl.Max = vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + panControls.Height + 550 - Me.Height
    scrControl.Left = Me.Width - scrControl.Width - 250
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 550
End Sub

Private Sub imgIcon_Click()
    Call txtPosCtr_DblClick
End Sub

'Private Sub lblBANDCode_Change()
'cmbBand = lblBANDCode
'End Sub

Private Sub lblLeadHand_Change()
    If lblLeadHand.Caption = "Y" Then
       chkLeadHand.Value = 1
    Else
       chkLeadHand.Value = 0
    End If
End Sub


Public Sub cmdCancel_Click()
    Dim X As Integer
    On Error GoTo Can_Err
    Dim PHMark As Variant
    
    'Data1.UpdateControls    ' returns without saving
    'data1.Recordset.CancelUpdate
    'If Not glbSQL and not glboracle Then Call Pause(0.5)
    'data1.Refresh
    fglbNew = False
    
    ''' Sam add July 2002 * Remove Binding Control
    rsDATA.CancelUpdate
    Call Display_Value
    
    'If Not rsDATA.EOF Then Call getCodes 'Ticket #28846 Franks 07/14/2016
    
    For X = 0 To 3 '2
        Call txtReptAuthority_Change(X)
    Next
    
    'Call ST_UPD_MODE(True)  ' reset screen's attributes
    'Call SET_UP_MODE
    fraPosition.Visible = False '26July99 js
    
    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        DoEvents
        If flgNewCancel And chkCurrent(0) Then
            Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
            
            'Update records with tracking on
            rsDATA("JH_TRK_CRS_RENEWAL") = False
            rsDATA.Update
            
            chkTrackCrsRenewal.Value = False
        End If
        flgNewCancel = False
        
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdJobFiles.Enabled = False
            
            'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
            If glbCompEntVacDaily Then
                cmdReCompDAccrual.Enabled = False
            End If
        Else
            cmdJobFiles.Enabled = True
            
            If Not gSec_Inq_Job_Files_Attachment Then
                cmdJobFiles.Enabled = False
            End If
        
            'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
            If glbCompEntVacDaily Then
                cmdReCompDAccrual.Enabled = True
            End If
        End If
    'End If
    
Exit Sub
    
Can_Err:
    If Err = 3018 Then
        Err = 0
        Resume Next
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_JOB_HISTORY", "Cancel")
    Call RollBack '26July99 js
End Sub

'Private Sub cmdCancel_GotFocus()

'Call SetPanHelp(Me.ActiveControl) '19Aug99 js

'If MDIMain.panHelp(0).Caption = " " Then
'    MDIMain.panHelp(0).Caption = "Save changes made"
'    MDIMain.panHelp(1).Caption = "Button"
'End If
'
'End Sub

Public Sub cmdClose_Click()
    Call NextForm
    Unload Me
    If glbOnTop = "FRMEPOSITION" Then glbOnTop = ""
End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub

Public Sub cmdDelete_Click()
    Dim a As Integer, Msg As String
    Dim SQLQ As String
    Dim xID
    Dim DeleteCurrentJob As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim UpdateAudit As Boolean
    Dim ODOA
    Dim xPayrollID As String
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
       MsgBox "Nothing to Delete"
       Exit Sub
    End If
    
    On Error GoTo Del_Err
    If glbVadim Then
        DeleteCurrentJob = False
        SQLQ = "Select JH_JOB,JH_EMPNBR,JH_ID FROM HR_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " "
        SQLQ = SQLQ & " AND JH_CURRENT <>0 "
        SQLQ = SQLQ & " AND JH_ID = " & Data1.Recordset!JH_ID
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTemp.EOF Then
            DeleteCurrentJob = True
        End If
        rsTemp.Close
        If DeleteCurrentJob Then
            If glbMulti Then
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID & " AND ED_PAYROLL_ID='" & txtPayrollID & "'"
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsTemp.EOF Then
                    MsgBox "The Current Position can not be deleted. Please enter the End Date instead"
                    rsTemp.Close
                    Exit Sub
                End If
                rsTemp.Close
            Else
                ODHRS = Val(medHours(0))
                oJob = clpJob
            End If
            
        End If
    Else
        oJob = clpJob
    End If
    ODOA = dlpStartDate
    
    Msg = "Are You Sure You Want To Delete This Record? "
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
       
    Screen.MousePointer = HOURGLASS

    '7.9 Enhancement - For all the clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        'Call procedure to delete the required courses of this position
        'Only if the position is current or tracked for course renewal
        If chkTrackCrsRenewal Or chkCurrent(0) Then
            If chkCurrent(0) Then
                Call Track_Courses_Renewal_Update("Delete", "C")
            Else
                Call Track_Courses_Renewal_Update("Delete", "P")
            End If
        End If
    'End If

    If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
        If Data1.Recordset("JH_CURRENT") <> 0 Then
            fOldFTE = Data1.Recordset("JH_FTENUM")
        Else
            fOldFTE = 0
        End If
    End If
    
    If glbMulti Then
        oENDDATE = dlpENDDATE.Text
        If oENDDATE <> "" Then
            If Not updFollow("D") Then
                Exit Sub
            End If
        End If
    End If
    
    xID = Data1.Recordset("JH_ID")
    If chkCurrent(0) Then DeleteCurrentJob = True
    
    'WDGPHU - Ticket #27899
    If glbCompSerial = "S/N - 2411W" And glbMulti And Not fglbNew Then
        'Current Position is being deleted, set the corresponding Salary's Current flag OFF
        If chkCurrent(0) Then
            Call SetCurrentSalary_OFF(glbLEE_ID, clpJob.Text, dlpStartDate.Text)
        End If
    End If
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_JOB_HISTORY WHERE JH_ID=" & xID
    gdbAdoIhr001.CommitTrans
    
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Delete from HRDOC_JOB_HISTORY where DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR = " & glbLEE_ID & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    
    Data1.Refresh
    
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        Call Set_Current_Flag
    Else
        Call Display_Value
    End If
    
    If glbVadim And DeleteCurrentJob Then
        If glbMulti Then
            Call DeletePayrollEmp(Date, glbLEE_ID, txtPayrollID.Text)
        Else
            UpdateAudit = False
            
            Dim HRChangs As New Collection
            If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
                If isChanged_Field(HRChangs, ODHRS, Data1.Recordset("JH_DHRS"), True) Then UpdateAudit = True
                If isChanged_Field(HRChangs, oJob, Data1.Recordset("JH_JOB")) Then UpdateAudit = True
            Else
                If isChanged_Field(HRChangs, ODHRS, medHours(0), True) Then UpdateAudit = True
                If isChanged_Field(HRChangs, oJob, clpJob) Then UpdateAudit = True
            End If
                
            If UpdateAudit = True Then
                Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID)
            End If
        End If
    End If
    
    fglbNew = False
    
    Call SET_UP_MODE
    
    If glbGuelph And (Not glbtermopen) Then
        Call AddFTE(glbLEE_ID, "DELE")
    End If
    'If Not glbMediPay Then
    '    Call Employee_Master_Integration(glbLEE_ID)
    'End If
    
    'Ticket #19687 - County of Lambton - Update Benefit records with correct Payroll ID of the Current job
    If glbVadim And glbLambton Then
        xPayrollID = Get_Payroll_ID_For_Benefit(glbLEE_ID)
        
        'Update employee's Benefit records with Payroll ID if Payroll ID found
        If xPayrollID <> "" Then
            SQLQ = "UPDATE HRBENFT SET BF_PAYROLL_ID = '" & xPayrollID & "' WHERE BF_EMPNBR = " & glbLEE_ID
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    
    If glbAdv Then 'Ticket #15282
        Call Employee_PositionDel_Integration(glbLEE_ID, oJob, ODOA, True)
    End If
    
    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        DoEvents
        
        'Track Courses for the previous Position which turned into Current
        If chkCurrent(0) And DeleteCurrentJob Then
            Call Display_Value
            Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        End If
    'End If
    
    'George commented on Mar 28,2006, In INFOHR, the user suppose to delete salary related the position
    'manually.
    'If glbCompSerial = "S/N - 2259W" Or glbGP Then 'Or (glbWFC And glbPlantCode = "GREN") Then
    '    Call Salary_Integration(glbLEE_ID, , True, False)
    'End If
    

    'Ticket #25911 Franks 12/17/2014 - begin
    'Release 8.1 - update the Budgeted Position
    If glbWFC Then
        If Len(clpJob.Text) > 0 Then
            Call mod_Upd_Pos_Budget_WFC(clpJob.Text, "")
        End If
        If Len(oJob) > 0 Then
            Call mod_Upd_Pos_Budget_WFC(oJob, "")
        End If
        Call InitData
    End If
    'Ticket #25911 Franks 12/17/2014 - end
    
    Screen.MousePointer = DEFAULT
    
Exit Sub
    
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_JOB_HISTORY", "Delete")
    Call RollBack '26July99 js
End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub

Public Sub cmdModify_Click()
Dim Response%, Msg$, Title$, DgDef As Double
Dim X% 'jaddy 10/25/99

On Error GoTo Mod_Err


If glbGuelph Then
    medFTENum.Enabled = False
    medFTEHrs.Enabled = False
End If

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    If chkCurrent(0).Value = True Then
        chkTrackCrsRenewal.Visible = False
    Else
        chkTrackCrsRenewal.Visible = True
    End If

    oTrkCrsRen = chkTrackCrsRenewal
End If

SavFte = medFTENum
SavFteHr = medFTEHrs
savCurrent = chkCurrent(0).Value

oENDDATE = dlpENDDATE.Text
oEndReason = clpCode(2).Text

For X% = 0 To 3 '2
    SavRpta(X%) = elpReptAuthShow(X%).Text
Next

savWHRS = medHours(1)
savSDate = dlpStartDate.Text
savJOB = clpJob.Text
fglbNew% = False
glbChgTermDate = ""
glbChgTermReason = ""
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    OBillingRate = medBillingRate
End If

'Kerry's Place - Ticket #24692 - Send Rehire Email for Rehired Employee
'Check if Rehired Employee
If glbCompSerial = "S/N - 2433W" Then
    If chkCurrent(0).Value = False And Len(Trim(dlpENDDATE.Text)) > 0 And Len(clpCode(2).Text) > 0 Then
        'No current position and there is a End Date and Reason. This happens from Rehired screen.
        flgRehire = True
    Else
        flgRehire = False
    End If
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    oPrimary = chkPrimary.Value
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_JOB_HISTORY", "Modify")
Call RollBack '26July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdNew_Click()
Dim SQLQ, Msg As String
Dim rsSal As New ADODB.Recordset
Dim xRes

On Error GoTo AddNP_Err

    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        If Not Data1.Recordset.EOF Then
        
            'Kerry's Place - Ticket #24692 - Send Rehire Email for Rehired Employee
            'Check if Rehired Employee
            Data1.Recordset.MoveFirst
            DoEvents
            If glbCompSerial = "S/N - 2433W" Then
                If chkCurrent(0).Value = False And Len(Trim(dlpENDDATE.Text)) > 0 And Len(clpCode(2).Text) > 0 Then
                    'No current position and there is a End Date and Reason. This happens from Rehired screen.
                    flgRehire = True
                Else
                    flgRehire = False
                End If
            End If
        
            'For Friensens only for now.
            'If chkCurrent(0).Value = False And glbCompSerial = "S/N - 2279W" Then 'And Not glbCompSerial = "S/N - 2188W" Then
            If glbCompSerial = "S/N - 2279W" Then 'And Not glbCompSerial = "S/N - 2188W" Then
                'Check first if End Date and End Reason had been entered for the older Position
                If Len(Trim(dlpENDDATE.Text)) = 0 Then
                    MsgBox "End Date cannot be left blank for this Position"
                    dlpENDDATE.SetFocus
                    Exit Sub
                ElseIf Not IsDate(dlpENDDATE.Text) Then
                    MsgBox "Invalid End Date"
                    dlpENDDATE.SetFocus
                    Exit Sub
                ElseIf CVDate(dlpENDDATE.Text) < CVDate(dlpStartDate.Text) Then
                    MsgBox "End Date cannot be prior to Start Date"
                    dlpENDDATE.SetFocus
                    Exit Sub
                ElseIf Len(Trim(clpCode(2).Text)) = 0 Then
                    MsgBox "End Reason cannot be left blank for this Position"
                    clpCode(2).SetFocus
                    Exit Sub
                ElseIf Not clpCode(2).ListChecker Then
                    MsgBox "Invalid End Reason"
                    clpCode(2).SetFocus
                    Exit Sub
                End If
            End If
            
            'City of Chatham Kent & 7.9 - Enhancement for all clients
            If glbCompSerial <> "S/N - 2279W" Then 'glbCompSerial = "S/N - 2188W" Then
                'Delete required courses
                Call Track_Courses_Renewal_Update("Delete", "C")
                
                'Update records with tracking on
                chkTrackCrsRenewal.Value = False
                rsDATA("JH_TRK_CRS_RENEWAL") = False
                rsDATA.Update
                
                GoTo Continue_NewClick
            Else
                'Confirm the Tracking Course Renewal ON or OFF
                Msg = "Do you want to track required courses renewals for this position? "
                xRes = MsgBox(Msg, vbYesNoCancel, "Confirm Required Course Renewal Tracking")
                If xRes = 7 Then    'No
                    'Delete required courses
                    Call Track_Courses_Renewal_Update("Delete", "C")
                    
                    'Update records with tracking on
                    chkTrackCrsRenewal.Value = False
                    rsDATA("JH_TRK_CRS_RENEWAL") = False
                    rsDATA.Update
                    
                    GoTo Continue_NewClick
                ElseIf xRes = 2 Then    'Cancel
                    Exit Sub
                Else
                    'Delete required courses ANYWAYS so that it can be added correctly with right type of position
                    'Changed
                    Call Track_Courses_Renewal_Update("Delete", "C")
                End If
                
                'Turn-ON tracking
                chkTrackCrsRenewal.Visible = True
                chkTrackCrsRenewal.Value = True
                
                'Call procedure to update/delete employee's Training list with this position's
                'required course list
                Call Track_Courses_Renewal_Update
                
                'Update records with tracking on
                rsDATA("JH_TRK_CRS_RENEWAL") = True
                rsDATA.Update
                
                chkTrackCrsRenewal.Value = False
                'chkTrackCrsRenewal.Visible = False
                DoEvents
            End If
            
        End If
    'End If
    
Continue_NewClick:
    fglbNew = True
    flgNewCancel = True
    
    If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
        fFTEDate = CurSDate
        fOldFTE = OFTE
    End If
    
    If glbLinamar Then
        clpJob.TransDiv = Right(glbLEE_ID, 3)
        clpCode(8).TransDiv = Right(glbLEE_ID, 3) 'Ticket #28846 Franks 07/14/2016
    End If
    If glbLambton Then
        chkUseForBenefit.Visible = True
    End If
    If glbWFC Then 'Ticket #27827 Franks 11/30/2015
        clpJob.TransDiv = glbWFCUserSecList
    End If

    Action = "A"
    
    'Ticket #29114 - WDGPHU - Employee History fix
    If glbCompSerial = "S/N - 2411W" Then
        medFTENum.Text = SavFte
    Else
        SavFte = ""
    End If
    SavFteHr = ""
    txtLambtonJob = ""
    fglbNew% = True
    
    If (Data1.Recordset.BOF Or Data1.Recordset.EOF) Then
        optSalary(1) = True
    Else
        ' "Not glbMulti" -> "Not glbMulti And Not glbSetPos" Modified by Frank 04/24/2001
        If Not glbMulti And Not glbSetPos Then
            fraPosition.Visible = True
            rsSal.Open "SELECT SH_EMPNBR FROM HR_SALARY_HISTORY WHERE SH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic
            If rsSal.EOF Then
                optSalary(1) = True
            Else
                optSalary(0) = True
            End If
            rsSal.Close
        End If
    End If
    
    Call SET_UP_MODE
    
    'George on Jan 26,2006 #10266
    If gsAttachment_DB Then
        glbJob = ""
        glbSDate = "01/01/1900"
        lblImport.Visible = True
        imgSec.Visible = False
        imgNoSec.Visible = True
        cmdImport.Visible = True
    End If
    'George on Jan 26,2006 #10266
    
    Call Set_Control("B", Me)
    
    If glbLinamar Then
        clpDiv = Right(glbLEE_ID, 3)
    End If
    
    'If fgetSection(lblEEID) = "GREN" Then
    If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab 'Ticket #14739
        If Len(locOrg) > 0 Then txtShift.Text = locOrg '"NS"
    End If
    If glbWFC Then
        txtShift.Text = "NS"
        'Ticket #24767 Franks 12/09/2013  - "   On a new record, repeat the Shift from the old record
        txtShift.Text = GetPrePositionData(glbLEE_ID, "JH_SHIFT", "NS")
        If glbAdv Then
            If NewHireForms.count > 0 And (glbUNION = "NONE" Or glbUNION = "EXEC") Then
            'Ticket #27605 Franks 10/13/2015
            '"   If the employee's union code is NONE or EXEC and the employee is a new hire
            'don't display the message about entering a Shift code for Tracker.
            Else
                MsgBox "For the Tracker integration, please ensure that the appropriate shift has been assigned to this employee."
            End If
        End If
        'Ticket #24767 Franks 12/09/2013 - end
        
        Call WFCDefaultHours

        If NewHireForms.count > 0 Then
            clpCode(1).Text = "NEW"
            Call WFCHRSoftDispValues 'Ticket #24184 Franks 09/11/2013
        Else
            'Ticket #25927 Franks 08/25/2014 - check if the HRSoft Position Upt Flag is YES
            If IsFirstEmpPosition(glbLEE_ID) Then
                If WFCHRSoftMissNewhire(glbLEE_ID, "SF_UPT_POSITION") Then
                    clpCode(1).Text = "NEW"
                    Call WFCHRSoftDispValues
                End If
            End If
        End If
        lblJobDesc = "" 'Ticket #27774 Franks 11/18/2015
        
        medFTENum.Text = 1 'Ticket #29005 Franks 08/02/2016
        If IsWFC_CONP Then  'Ticket #30358 Franks 07/13/2017 - leave these as blank for Independent Contractor Positions
            medFTEHrs.Text = 0
        Else
            medFTEHrs.Text = 2080 'Ticket #29005 Franks 08/02/2016
        End If
    End If '--------------- WFC end
    
    If glbCompSerial = "S/N - 2379W" Then 'Town of LaSalle Ticket #14534
        txtShift.Text = "NOSD"
    End If
    If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #21528 Franks 02/02/2012
        txtShift.Text = "0"
    End If
    
    rsDATA.AddNew
    
    chkCurrent(0) = glbMulti
    lblEEID = glbLEE_ID
    lblCompNo.Caption = "001"
    
    Call SetDefaultValue
    
    If glbCompSerial = "S/N - 2241W" Then 'Granite Club
        If NewHireForms.count > 0 Then 'New Hire only
            chkActPosition.Value = True
        End If
    End If
    
    'Municipality of North Perth  Ticket #19209 Franks 05/09/2011
    If glbCompSerial = "S/N - 2429W" Then
        chkActPosition.Value = True
    End If
    
    If NewHireForms.count > 0 Then 'From v7.6
        dlpStartDate = GetDoh(glbLEE_ID)
    End If
    
    If IsWFC_CONP Then 'Ticket #30359 Franks 07/11/2017
        clpJob.Enabled = False
        Call WFC_CONP_Fields
    Else
        clpJob.Enabled = True
        clpJob.SetFocus
    End If
    
     'Simona - begin - Assessment Strategies-#14963
    If (glbCompSerial = "S/N - 2401W") Then
        medHours(0).Text = "7.5"
        medHours(1).Text = "37.5"
        medHours(2).Text = "75.0"
        medFTENum.Text = "1"
        medFTEHrs.Text = "1950"
    End If
    'Simona - end - Assessment Strategies-#14963

    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        cmdJobFiles.Enabled = False
        chkTrackCrsRenewal.Enabled = False
        chkTrackCrsRenewal.Visible = False
    'End If

    If glbCompSerial = "S/N - 2259W" Then 'Ticket #16877
        If Not glbMulti Then
            chkActPosition.Value = True
        End If
        chkCurrent(0).Value = True 'Ticket #21599
    End If

    If glbCompSerial = "S/N - 2382W" Then 'Ticket #18755 Samuel
        ''When clicking a new record, default Reporting Authorities (1 – 3) from previous record
        'Ticket #21652 Franks 03/20/2012 comment out this function
        'Use SAM_POS_ITEMS_MATRIX to setup these
        'Call GetPreReportAuthorities(glbLEE_ID)
        
        'Ticket #20371 Franks 05/25/2011 - get the current salary effective date
        dlpCurSEDate.Text = GetCurSalEDate(glbLEE_ID)
    End If
    
    'Ticket #20105 Franks 09/20/2011
    If glbCompSerial = "S/N - 2384W" Then 'Town of St. Marys
        chkActPosition.Value = True
    End If
    
    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Enabled = False
    End If
    
    If glbWFC Then Call WFC_fraPosition 'Ticket #29343 Franks 10/18/2016

Exit Sub

AddNP_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_JOB_HISTORY", "Add")
Call RollBack '26July99 js

End Sub

Private Function GetCurSalEDate(xEmpNo) 'Samuel Ticket #20371 Franks 05/25/2011
Dim rsCurSal As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT SH_EMPNBR, SH_EDATE FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNo & " AND NOT (SH_CURRENT = 0) "
    SQLQ = SQLQ & "ORDER BY SH_EDATE DESC "
    If rsCurSal.State <> 0 Then rsCurSal.Close
    rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsCurSal.EOF Then
        If Not IsNull(rsCurSal("SH_EDATE")) Then
            retVal = rsCurSal("SH_EDATE")
        End If
    End If
    rsCurSal.Close
    GetCurSalEDate = retVal
End Function

Private Function GetPrePositionData(xEmpNo, xFieldName, xDefault) 'Ticket #24767 Franks 12/09/2013
Dim rsPreEJob As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = xDefault
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " AND NOT (JH_CURRENT = 0) "
    SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
    If rsPreEJob.State <> 0 Then rsPreEJob.Close
    rsPreEJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPreEJob.EOF Then
        If Not IsNull(rsPreEJob(xFieldName)) Then
            retVal = rsPreEJob(xFieldName)
        End If
    End If
    rsPreEJob.Close
    GetPrePositionData = retVal
End Function

Private Sub GetPreReportAuthorities(xEmpNo)
Dim rsPreEJob As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " AND NOT (JH_CURRENT = 0) "
    SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
    If rsPreEJob.State <> 0 Then rsPreEJob.Close
    rsPreEJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPreEJob.EOF Then
        If Not IsNull(rsPreEJob("JH_REPTAU")) Then
            elpReptAuthShow(0).Text = IIf(Not IsNull(rsPreEJob("JH_REPTAU")), rsPreEJob("JH_REPTAU"), "")
        End If
        If Not IsNull(rsPreEJob("JH_REPTAU2")) Then
            elpReptAuthShow(1).Text = IIf(Not IsNull(rsPreEJob("JH_REPTAU2")), rsPreEJob("JH_REPTAU2"), "")
        End If
        If Not IsNull(rsPreEJob("JH_REPTAU3")) Then
            elpReptAuthShow(2).Text = IIf(Not IsNull(rsPreEJob("JH_REPTAU3")), rsPreEJob("JH_REPTAU3"), "")
        End If
        'Ticket #20052 Franks 07/15/2011 - begin
        If Not IsNull(rsPreEJob("JH_REPTAU4")) Then
            elpReptAuthShow(3).Text = IIf(Not IsNull(rsPreEJob("JH_REPTAU4")), rsPreEJob("JH_REPTAU4"), "")
        End If
        If Not IsNull(rsPreEJob("JH_PROFIT_SHARING")) Then
            If rsPreEJob("JH_PROFIT_SHARING") Then
                chkProSha.Value = 1
            Else
                chkProSha.Value = 0
            End If
        End If
        'Ticket #20052 Franks 07/15/2011 - end
    End If
    rsPreEJob.Close
End Sub

Public Sub cmdOK_Click()
Dim X%, xID, xFte, xFteHr
Dim rsJOB As New ADODB.Recordset
Dim rsJOBMASTER As New ADODB.Recordset
Dim xReptAuthority, xChange
Dim SQLQ, Msg, startDate
Dim rs As New ADODB.Recordset
Dim oCurrent As Boolean
Dim xPayrollID As String
Dim xBranch  As String

On Error GoTo Add_Err

'Ticket #23773 - Do not need this functionality any more
'City of Timmins - Ticket #13207
'If glbCompSerial = "S/N - 2375W" And fglbNew <> True Then
'    'Check if End Date entered then do not prompt for Password
'    If oENDDATE = "" And dlpENDDATE.Text <> "" And chkCurrent(0) = False Then
'        'Save the changes and do not prompt for Password
'    Else
'        'Ask for the password
'        glbAccessPswd = False
'        frmAccessPswd.Show 1
'        If glbAccessPswd = False Then   'Access Denied
'            Call cmdCancel_Click
'            Exit Sub
'        End If
'    End If
'End If

If glbVadim And glbMulti And Not fglbNew Then
    If Not chkVadimPayrollID Then Exit Sub
End If

If Not chkPosition() Then Exit Sub

Dim xRes As Integer

'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    If flgTrainLstReset = False Then
        'Check first if the Position Start Date has changed for Current or Tracked positions
        If Not fglbNew Then
            If dlpStartDate.Text <> savSDate And (chkTrackCrsRenewal Or chkCurrent(0)) Then
                'Call proceedure to update Training List record with new Start Date and
                'also recalculate renewal date for Courses Not Taken as the Renewal Dates were
                'based on Position Start Date
                Call Update_Position_Start_Date_in_Training_List(savSDate, dlpStartDate.Text)
            End If
        End If
    
        If chkTrackCrsRenewal.Visible = True Then
            If oTrkCrsRen <> chkTrackCrsRenewal And chkCurrent(0).Value = False Then
                'Confirm the Tracking Course Renewal ON or OFF
                If chkTrackCrsRenewal Then
                    Msg = "Are you sure you want to turn-ON the Required Courses Renewals? "
                Else
                    Msg = "Are you sure you want to turn-OFF the Required Courses Renewals? "
                End If
                xRes = MsgBox(Msg, 36, "Confirm Required Course Renewal Tracking")
                If xRes <> 6 Then   'No
                    If chkTrackCrsRenewal Then
                        chkTrackCrsRenewal.Value = False    'undo the checking
                    Else
                        chkTrackCrsRenewal.Value = True     'undo the checking
                    End If
                    Exit Sub
                End If
                
                'Call procedure to update/delete employee's Training list with this position's
                'required course list
                Call Track_Courses_Renewal_Update
            End If
            
            'Hold current value of the Current flag
            oCurrent = chkCurrent(0).Value
        Else
            'City of Chatham-Kent - Ticket #16794
            If glbCompSerial = "S/N - 2188W" Then
                'Call procedure to delete employee's Training list with this position's
                'required course list
                'Call Track_Courses_Renewal_Update
            End If
            
            'Hold current value of the Current flag
            oCurrent = chkCurrent(0).Value
        End If
    Else
        'Hold current value of the Current flag
        oCurrent = chkCurrent(0).Value
    End If
'End If

If Not glbSetPos Then Call UpdPositionCCAC

Screen.MousePointer = HOURGLASS

Call UpdUStats(Me) ' update user's stats (who did it and when)

If glbCompSerial = "S/N - 2259W" And (Not glbtermopen) Then
    SQLQ = "SELECT ED_ORG, ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("ED_SECTION") <> "Y" Then
            If IsNull(rs("ED_ORG")) = False Then clpCode(0).Text = rs("ED_ORG")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If

'City of Pickering - Ticket #13281
If glbCompSerial = "S/N - 2217W" Then
    If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
        If IsNumeric(medHours(2)) And (medFTEHrs.Text = "" Or medFTEHrs.Text = "0") Then     'Hours/Pay Period
            medFTEHrs = medHours(2) * 26
        End If
    End If
End If

If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
    If fglbNew Then
        If IsNumeric(medFTENum) Then
            fNewFTE = Val(medFTENum)
        Else
            fNewFTE = 0
        End If
    End If
End If

If glbSamuel Then 'Ticket #20885 Franks 12/01/2011
    Call SAMUEL_Trans(glbLEE_ID)
End If

'Wellington-Dufferin-Guelph Public Health - Ticket #22635
If glbCompSerial = "S/N - 2411W" Then
    If fglbNew Then
        If Len(elpReptAuthShow(1).Text) = 0 Then
            elpReptAuthShow(1).Text = elpReptAuthShow(0).Text
        End If
    End If
End If

Call UpdCodes 'Ticket #28846 Franks 07/14/2016

Call Set_Control("U", Me, rsDATA)

For X = 0 To 2
    xReptAuthority = getEmpnbr(elpReptAuthShow(X))
    rsDATA("JH_REPTAU" & IIf(X = 0, "", X + 1)) = IIf(Val(xReptAuthority) = 0, Null, xReptAuthority)
Next
If glbLinamar Then
    'If chkActPosition Then
    '    rsDATA("JH_POSITION_CONTROL") = "YES"
    'Else
    '    rsDATA("JH_POSITION_CONTROL") = "NO"
    'End If
ElseIf glbMulti Then 'George on Dec 7,2005 #9928 begin
    If chkActPosition Then
        If fglbNew Then
            xID = 0
        Else
            xID = rsDATA!JH_ID
        End If
        SQLQ = "UPDATE HR_JOB_HISTORY"
        SQLQ = SQLQ & " SET JH_POSITION_CONTROL = 'NO' "
        SQLQ = SQLQ & " WHERE JH_EMPNBR =" & glbLEE_ID & " AND JH_ID <> " & xID
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
        'rsDATA("JH_POSITION_CONTROL") = "YES"
    Else
        'rsDATA("JH_POSITION_CONTROL") = "NO"
    End If 'George on Dec 7,2005 #9928 end
End If

'Hemu - Ottawa CCAC uses this field to record their own CCAC Position # and so it cannot be
'set to NO or YES - Ticket #11411
If Not glbOttawaCCAC Then
    If chkActPosition Then
        rsDATA("JH_POSITION_CONTROL") = "YES"
    Else
        rsDATA("JH_POSITION_CONTROL") = "NO"
    End If
End If

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    'gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    xID = rsDATA!JH_ID
    'gdbAdoIhr001X.CommitTrans
    'rsDATA.Resync
    'George Jan 26,2006
    If gsAttachment_DB Then
        'gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_JOB_HISTORY set DJ_JOB='" & rsDATA("JH_JOB") & "',DJ_SDATE=" & Date_SQL(rsDATA("JH_SDATE")) & " where DJ_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
        'gdbAdoIhr001_DOC.CommitTrans
        'glbJob = rsDATA!JH_JOB
        'glbSDate = rsDATA!JH_SDATE
    End If
    'George Jan 26,2006
Else
    'gdbAdoIhr001.BeginTrans
    rsDATA.Update
    xID = rsDATA!JH_ID
    'gdbAdoIhr001.CommitTrans
    'rsDATA.Requery
    'xID = rsDATA!JH_ID
    'George Jan 26,2006
    If gsAttachment_DB Then
        'gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update HRDOC_JOB_HISTORY set DJ_JOB='" & rsDATA("JH_JOB") & "',DJ_SDATE=" & Date_SQL(rsDATA("JH_SDATE")) & " where DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR = " & glbLEE_ID & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
        'gdbAdoIhr001_DOC.CommitTrans
        'glbJob = rsDATA!JH_JOB
        'glbSDate = rsDATA!JH_SDATE
    End If
    'George Jan 26,2006
End If

'Add by Franks on Jul 11,02 for ticket #2546
'If glbWFC And lblBANDCode.Visible Then
If glbWFC And clpCode(6).Visible Then
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & clpJob.Text & "' "
    rsJOBMASTER.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJOBMASTER.EOF Then
        If rsJOBMASTER("JB_BAND") <> clpCode(6).Text Then
            rsJOBMASTER("JB_BAND") = clpCode(6).Text
            rsJOBMASTER.Update
            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_BAND = '" & clpCode(6).Text & "' "
            SQLQ = SQLQ & "WHERE SH_JOB = '" & clpJob.Text & "' "
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute SQLQ
            gdbAdoIhr001.CommitTrans
        End If
    End If
    rsJOBMASTER.Close
End If
'Add by Franks on Jul 11,02 for ticket #2546

If glbWFC Then 'Ticket #29438 Franks 11/08/2016 - not new hire, new position, salaried employee
    If xIsWFCRetpEmpShowUp Then
        If IsWFCReptAuth(glbLEE_ID, "") Then
            glbWFC_IPPopFormName = "WFCEmpListWithRept"
            glbWFC_IncePlanID = glbLEE_ID
            frmCheckListView.lblStDate = dlpStartDate.Text
            frmCheckListView.Show 1
        End If
    End If
End If

'Frank 12/16/2009
'WFC Pension Outstanding Tasks By Dec1009.doc
If glbWFC Then
    If fglbNew Then 'new record only
        If WFCPensionEligible(glbLEE_ID) Then
            Call WFCPensionMasUpt(glbLEE_ID, "Position_NOGC", dlpStartDate, clpJob.Text, Year(dlpStartDate))
        End If
        'Ticket #22613 Franks
        Call WFCPosSkillsUpd(glbLEE_ID, clpJob.Text, dlpStartDate.Text)
    End If
End If
'Data1.Refresh

'Burlington Tech Ticket #13235
'If the new position code is found in Backup Position table, delete it.
If glbCompSerial = "S/N - 2351W" Then
    If fglbNew And (Not glbtermopen) Then
        SQLQ = "DELETE FROM HR_JOB_BACKUP WHERE JH_EMPNBR = " & glbLEE_ID & " "
        SQLQ = SQLQ & "AND JH_JOB = '" & clpJob.Text & "' "
        gdbAdoIhr001.Execute SQLQ
    End If
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti And Not fglbNew Then
    'Current Position is turned OFF, set the corresponding Salary's Current flag OFF as well
    If oENDDATE = "" And IsDate(dlpENDDATE.Text) And chkCurrent(0) = False Then  'And clpCode(2).Text <> "" Then
        Call SetCurrentSalary_OFF(glbLEE_ID, clpJob.Text, dlpStartDate.Text)
    End If
End If

Call Set_Current_Flag
Data1.Refresh
Data1.Recordset.Find "JH_ID=" & xID

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            'glbJob = xID
            glbJob = Data1.Recordset("JH_JOB")
            glbSDate = Data1.Recordset("JH_SDATE")
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
        End If
    End If
    glbDocImpFile = ""
End If

chkCurrent(0) = Data1.Recordset("JH_CURRENT")

'Ticket #19687 - County of Lambton - Update Benefit records with Payroll ID (new)
If glbVadim And glbLambton Then
    xPayrollID = Get_Payroll_ID_For_Benefit(glbLEE_ID)
    
    'Update employee's Benefit records with Payroll ID if Payroll ID found
    If xPayrollID <> "" Then
        SQLQ = "UPDATE HRBENFT SET BF_PAYROLL_ID = '" & xPayrollID & "' WHERE BF_EMPNBR = " & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
    End If
End If


If gsEMAIL_ONPOSITION Then 'Ticket #21444 Franks 02/10/2012
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
        MailBody = getWaltersIncEmailBody
    Else
        MailBody = ""
        If NewHireForms.count = 0 Then 'Non new hire
            'savJOB
            'If Len(OJOB) > 0 And Not OJOB = clpJob.Text And chkCurrent(0) Then
            If Len(savJOB) > 0 And Not savJOB = clpJob.Text And chkCurrent(0) Then
                MailBody = MailBody & "This will serve to confirm that the following employee's position title has been changed" & vbCrLf & vbCrLf
                'Len(OJOB) > 0 - not the first job, only for change
                ' chkCurrent(0) - current position only
                If glbCompSerial = "S/N - 2382W" Then  'Samuel
                    'MailBody = MailBody & GetEmailBodyForSamuel(glbLEE_ID) & vbCrLf
                    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf & vbCrLf
                    xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
                    If Len(xBranch) > 0 Then
                        xBranch = GetTABLDesc("EDSE", xBranch)
                    End If
                    MailBody = MailBody & "Branch: " & xBranch & vbCrLf & vbCrLf
                Else
                    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf & vbCrLf
                End If
                MailBody = MailBody & "Previous position: " & getPosDesc(savJOB) & vbCrLf & vbCrLf
                MailBody = MailBody & "New position: " & getPosDesc(clpJob.Text) & vbCrLf & vbCrLf
                MailBody = MailBody & "Effective Date: " & dlpStartDate.Text & vbCrLf
                If glbWFC Then 'Ticket #29183 Franks 09/12/2016
                    MailBody = MailBody & vbCrLf & lblReptAuth(0).Caption & ": " & elpReptAuthShow(0).Caption & vbCrLf
                    If Not xRept1FromDivMaster = txtReptAuthority(0).Text Then
                        If Len(xRept1FromDivMaster) > 0 Then
                            MailBody = MailBody & lblReptAuth(0).Caption & " " & xRept1FromDivMaster & "(" & GetEmpData(xRept1FromDivMaster, "ED_SURNAME") & "," & GetEmpData(xRept1FromDivMaster, "ED_FNAME") & ")" & " was changed to " & txtReptAuthority(0).Text & "(" & elpReptAuthShow(0).Caption & "). This does not match the Reporting Authority as setup in the Position Master." & vbCrLf
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Call procedure to add required courses to the Training List
    If fglbNew And chkCurrent(0) Then
        Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
    Else
        'Position Code has changed. Delete the Training List of the older Position and then create new
        'Training List for the changed Position Code. (Ticket #22044)
        If Len(savJOB) > 0 And Not savJOB = clpJob.Text And chkCurrent(0) Then
            Call Track_Courses_Renewal_Update("Delete", "C", savJOB)
        End If
        
        'If the Current value had changed and is Current
        If (chkCurrent(0) And oCurrent <> chkCurrent(0)) Or (savCurrent = False And chkCurrent(0)) Or (Len(savJOB) > 0 And Not savJOB = clpJob.Text And chkCurrent(0)) Then
            Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        End If
    End If
'End If

If chkCurrent(0) Then
    SQLQ = ""
    For X = 0 To 2
        xReptAuthority = getEmpnbr(elpReptAuthShow(X))
        If SavRpta(X%) <> elpReptAuthShow(X%).Text Then xChange = True
        SQLQ = SQLQ & " PH_REPTAU" & IIf(X = 0, "", X + 1) & " =" & IIf(Val(xReptAuthority) > 0, xReptAuthority, "Null") & IIf(X = 2, " ", ",")
    Next
    If Action = "M" Then
        If savJOB <> clpJob.Text Or savSDate <> dlpStartDate.Text Then
        Else
            If xChange Then
                SQLQ = "UPDATE HR_PERFORM_HISTORY SET " & SQLQ
                SQLQ = SQLQ & " WHERE PH_EMPNBR=" & glbLEE_ID & " AND PH_JOB='" & clpJob.Text & "' AND PH_CURRENT<>0 "
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
        If savWHRS <> medHours(1) Then
            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_WHRS =" & Val(medHours(1))
            SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_JOB='" & clpJob.Text & "' AND SH_CURRENT<>0 "
            gdbAdoIhr001.Execute SQLQ
            
            'Hemu
            savWHRS = medHours(1)
            'Hemu
            
        End If
        If savGrid <> clpGrid.Text And glbMultiGrid Then
            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_GRID ='" & clpGrid.Text & "'"
            SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_JOB='" & clpJob.Text & "' AND SH_CURRENT<>0 "
            gdbAdoIhr001.Execute SQLQ
            savGrid = clpGrid.Text
        End If
        
        'End If
    End If
    If Not glbMulti Then
        SQLQ = "UPDATE HREMP SET ED_SHIFT ='" & txtShift & "' WHERE ED_EMPNBR=" & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
    End If
    If Val(medHours(0)) <> Val(fgtxtDhrs) Then
        SQLQ = "UPDATE HREMP SET ED_DHRS =" & Val(medHours(0)) & " WHERE ED_EMPNBR=" & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRENTHRS SET HE_DHRS =" & Val(medHours(0)) & " WHERE HE_EMPNBR=" & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
        glbENTScreen = True
    End If
    If SavFte <> medFTENum Or SavFteHr <> medFTEHrs Then
        If SavFte <> medFTENum Then xFte = SavFte Else xFte = ""
        If SavFteHr <> medFTEHrs Then xFteHr = SavFteHr Else xFteHr = ""
        If Not EmpHisCalc(3, glbLEE_ID, "", "", "", "", "", xFte, xFteHr, Date) Then MsgBox "EMPHIS Error"
    End If
    'Ticket #27553 Franks 09/21/2015 - begin
    If Not savJOB = clpJob.Text And chkCurrent(0) Then
        'Ticket #29722 - When testing this for Multi Position client - added a new position - it updated with blank values for both old and new values
        'If Not EmpHisCalc(7, glbLEE_ID, "", "", "", "", "", "", "", Date, , savJOB) Then MsgBox "EMPHIS Error "
        If Not EmpHisCalc(7, glbLEE_ID, "", "", "", "", "", "", "", Date, , clpJob.Text, , , , , savJOB) Then MsgBox "EMPHIS Error "
    End If
    If Not oREPTAU = txtReptAuthority(0).Text And chkCurrent(0) Then
        'Ticket #29722 - When testing this for Multi Position client - added RA 1 - it updated with blank values for both old and new values
        'If Not EmpHisCalc(8, glbLEE_ID, "", "", "", "", "", "", "", Date, , oREPTAU) Then MsgBox "EMPHIS Error "
        If Not EmpHisCalc(8, glbLEE_ID, "", "", "", "", "", "", "", Date, , txtReptAuthority(0).Text, , , , , oREPTAU) Then MsgBox "EMPHIS Error "
    End If
    'Ticket #27553 Franks 09/21/2015 - end
    If Not glbMulti Then
        If clpJob.Text <> fgtxtjob Or fgtxtStartDate <> CVDate(dlpStartDate.Text) Then
            If optSalary(0).Value = True Then
                Call Upd_Related_Salary
            End If
        End If
    End If
    If Not glbMulti Then
        'Hemu 07/02/2003 Begin - Ticket #4247, Update Employment Equity Data with NOC Code
        Dim rsEmpNOC As New ADODB.Recordset
        
        rsEmpNOC.Open "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID & " AND JH_CURRENT <> 0)", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpNOC.EOF Then
            If Not IsNull(rsEmpNOC("JB_FEDGRP")) Then
                gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NOGC = '" & rsEmpNOC("JB_FEDGRP") & "' WHERE EQ_EMPNBR = " & glbLEE_ID
            End If
        End If
        rsEmpNOC.Close
        'Hemu 07/02/2003 End - Ticket #4247
    End If
    
    'WDGPHU - Ticket #27899
    If glbCompSerial = "S/N - 2411W" And glbMulti Then
        'Update employee's Salary records with the correct Primary Position checkbox
        Call UpdatePrimaryPositionSalary(glbLEE_ID)
    End If
    
    Call InitData
End If

If glbCompSerial = "S/N - 2217W" Then ' FOR CITY OF PICKERING
    If chkCurrent(0) Then
        If Not updFollow("U") Then Exit Sub
    End If
End If

'Macaulay Child Dev. 8.0 - Ticket #24564: Create a Follow Up record when Current, End Date entered and no End Reason
If glbCompSerial = "S/N - 2420W" Then
    If chkCurrent(0) And IsDate(dlpENDDATE.Text) And clpCode(2).Text = "" Then
        If Not updFollow("U") Then Exit Sub
    ElseIf chkCurrent(0) And dlpENDDATE.Text = "" And clpCode(2).Text = "" Then
        If Not updFollow("U") Then Exit Sub
    End If
End If

If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
    If fglbNew Then
        Call Pause(0.5)
        Call AddFTE(glbLEE_ID, "NEW")
    End If
End If

If glbOttawaCCAC Then
   If chkCurrent(0) Then
        Call UpdOttawaCCAC
   End If
End If

If glbCompSerial = "S/N - 2347W" Then
    Call updBenefitForSurreyPlace(glbLEE_ID)
End If

If Not glbMediPay Then
    Call Employee_Master_Integration(glbLEE_ID)
End If

'George Mar 9 2006 commented. Moved to Upd_Related_Salary. Here could not know this position changed will create a new salary record in Salary_History or just change.
'If glbCompSerial = "S/N - 2259W" Or glbGP Then
'    Call Salary_Integration(glbLEE_ID, , False, IIf(fglbNew% = 0, False, True))
'End If
'aded by Bryan 22/09/05 Ticket# 9368
'Ticket #17130 Frank 08/04/2009
If glbGP And (fglbNew% = 0) Then 'Not new record
    Call Salary_Integration(glbLEE_ID, , False, IIf(fglbNew% = 0, False, True))
End If

If glbMediPay Then 'Ticket #14752
    'Hemu - Ticket #14752 - Because Job Start Date and Reason for Change needs to be passed
    'as well as Salary Effective Date and Reason for Change whenever these happens, I had to
    'pass to separate the function out.
    'Call Salary_Integration(glbLEE_ID)
    Call Position_Integration(glbLEE_ID)
End If

If NewHireForms.count > 0 And glbCompSerial = "S/N - 2375W" Then
    Call updateOMERS
End If

'added by Bryan 12/Apr/06 Ticket#10644
If isEDU Then
    If elpReptAuthShow(0).Text <> "" Then
        If glbCompSerial = "S/N - 2347W" And NewHireForms.count > 0 Then 'Surreyplace
            
            SQLQ = "SELECT HRE_SCHEDULE.SC_CLASSID, HRE_SCHEDULE.SC_DATE FROM HRE_COURSE INNER JOIN HRE_SCHEDULE ON HRE_COURSE.CS_ID = HRE_SCHEDULE.SC_CLASSID "
            SQLQ = SQLQ & "WHERE HRE_SCHEDULE.SC_DATE > " & Date_SQL(Date) & " AND HRE_COURSE.CS_CODE='ORIE' ORDER BY HRE_SCHEDULE.SC_DATE ASC"
            rs.Open SQLQ, gdbAdoIHREDU, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                SQLQ = "INSERT INTO HRE_ENROLLMENT(EN_EMPNBR, EN_TYPE, EN_CLASSID, EN_WAITING, EN_NAME, EN_SUPER, EN_LDATE, EN_LUSER, EN_LTIME) "
                SQLQ = SQLQ & "VALUES (" & glbLEE_ID & ", 'E', " & rs("SC_CLASSID") & ", 1, '" & Replace(lblEEName.Caption, "'", "") & "'," & elpReptAuthShow(0).Text & ", "
                SQLQ = SQLQ & Updstats(0) & ", '" & Updstats(2) & "', '" & Updstats(1) & "')"
                gdbAdoIHREDU.BeginTrans
                gdbAdoIHREDU.Execute SQLQ
                gdbAdoIHREDU.CommitTrans
            End If
            rs.Close
        End If
        
        If glbCompSerial = "S/N - 2347W" And fglbNew Then 'Surreyplace
            Dim xclass As String
            
            SQLQ = "SELECT HRE_SCHEDULE.SC_CLASSID, HRE_COURSE.CS_CODE, HRE_SCHEDULE.SC_DATE FROM HRE_COURSE INNER JOIN HRE_SCHEDULE ON HRE_COURSE.CS_ID = HRE_SCHEDULE.SC_CLASSID "
            SQLQ = SQLQ & "WHERE HRE_SCHEDULE.SC_DATE > " & Date_SQL(Date) & " AND HRE_COURSE.CS_CODE IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB='" & clpJob.Text & "')"
            SQLQ = SQLQ & "ORDER BY HRE_SCHEDULE.SC_DATE ASC"
            rs.Open SQLQ, gdbAdoIHREDU, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                xclass = ""
                Do
                    If xclass <> rs("CS_CODE") Then
                        xclass = rs("CS_CODE")
                        SQLQ = "INSERT INTO HRE_ENROLLMENT(EN_EMPNBR, EN_TYPE, EN_CLASSID, EN_WAITING, EN_NAME, EN_SUPER, EN_LDATE, EN_LUSER, EN_LTIME) "
                        SQLQ = SQLQ & "VALUES (" & glbLEE_ID & ", 'E', " & rs("SC_CLASSID") & ", 1, '" & Replace(lblEEName.Caption, "'", "") & "'," & elpReptAuthShow(0).Text & ", "
                        SQLQ = SQLQ & Updstats(0) & ", '" & Updstats(2) & "', '" & Updstats(1) & "')"
                        gdbAdoIHREDU.BeginTrans
                        gdbAdoIHREDU.Execute SQLQ
                        gdbAdoIHREDU.CommitTrans
                    End If
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            rs.Close
        End If
    End If
Set rs = Nothing
End If
'end Bryan

'Linamar - Post a message on Save click for a New Position - Ticket #17409
If fglbNew And glbLinamar Then
    MsgBox "Verify require Training.", vbOKOnly, "info:HR - Verify Training"
End If

'Ticket #18790 - Update EEO record
If glbEmpCountry = "U.S.A." Then
    If fglbNew Then
        'Ticket #20852
        'Getting error for multi position employee because there is more than one Position.
        'The EO_OCC_CAT that needs to be updated based on the Current Position from the HRJOB
        'results in more than one record, hence error. So passing the position that is being added
        'and that will fix the error.
        If glbMulti Then
            Call uptEEO_Fields(glbLEE_ID, "Update", , , clpJob.Text)
        Else
            Call uptEEO_Fields(glbLEE_ID, "Update")
        End If
    End If
End If

If glbWFC Then 'Ticket #23117 Franks 01/28/2013
    If NewHireForms.count > 0 Then
        If fglbNew Then 'new record only
            'Ticket #23575 Franks 04/12/2013 - Remove from program
            'Call WFC_PT_PenCheck("Y")
        End If
    End If
    'Ticket #25911 Franks 12/17/2014 - begin
    'Release 8.1 - update the Budgeted Position
    'If fglbNew Then
        'Ticket #28341 Franks 03/21/2016 - user may change FTE Hours/Year only
        If (Not savJOB = clpJob.Text) Or (Not SavFte = medFTENum.Text) Or (Not SavFteHr = medFTEHrs.Text) Then
            Call mod_Upd_Pos_Budget_WFC(clpJob.Text, "")
            If Not savJOB = clpJob.Text Then
                If Len(savJOB) > 0 Then
                    Call mod_Upd_Pos_Budget_WFC(savJOB, "")
                End If
            End If
        End If
    'End If
    'Ticket #25911 Franks 12/17/2014 - end
End If

If glbSamuel Then 'Ticket #23386 Franks 03/25/2013
    If NewHireForms.count > 0 Then
        Call Samuel_Vac_Ent_Cal(glbLEE_ID)
    End If
End If

If glbWFC Then 'Ticket #28763 Franks 06/21/2016
    If NewHireForms.count > 0 And gsEMAIL_ONNEWHIRE Then 'Send email on New Hire
        Call WFCNewHireEmailSending
    End If
End If

'Kerry's Place - Ticket #24692 - Send New/Rehire Emails
If glbCompSerial = "S/N - 2433W" Then
    'Send email on New Hire
    If NewHireForms.count > 0 And gsEMAIL_ONNEWHIRE Then
        MailBodyN = "The new employee has been hired." & vbCrLf & vbCrLf
        MailBodyN = MailBodyN & "Employee #: " & glbLEE_ID & vbCrLf
        MailBodyN = MailBodyN & "Start Date: " & GetEmpData(glbLEE_ID, "ED_DOH") & vbCrLf
        MailBodyN = MailBodyN & "First Name: " & GetEmpData(glbLEE_ID, "ED_FNAME") & vbCrLf
        MailBodyN = MailBodyN & "Last Name: " & GetEmpData(glbLEE_ID, "ED_SURNAME") & vbCrLf
        
        MailBodyN = MailBodyN & "Position: " & GetJobData(clpJob.Text, "JB_DESCR") & vbCrLf
        MailBodyN = MailBodyN & lStr("Rept. Authority 1") & ": " & GetEmpData(txtReptAuthority(0).Text, "ED_SURNAME") & ", " & GetEmpData(txtReptAuthority(0).Text, "ED_FNAME") & vbCrLf
        MailBodyN = MailBodyN & lStr("Union") & ": " & GetTABLDesc("EDOR", clpCode(0).Text) & vbCrLf
        MailBodyN = MailBodyN & lStr("Location") & ": " & GetTABLDesc("EDLC", GetEmpData(glbLEE_ID, "ED_LOC")) & vbCrLf
        MailBodyN = MailBodyN & lStr("Department") & ": " & GetDeptName(clpDept.Text, "DF_NAME") & vbCrLf
        MailBodyN = MailBodyN & lStr("Division") & ": " & Get_Division_Name(clpDiv.Text, "Division_Name") & vbCrLf
        'Screen.MousePointer = DEFAULT
        Call imgEmail_NewHire
        'Screen.MousePointer = HOURGLASS
    End If
    
    'Send Rehire Email for Rehired Employee
    If NewHireForms.count = 0 And gsEMAIL_ONREHIRE And flgRehire And chkCurrent(0).Value = True Then  'Non new hire
        MailBodyR = "This employee has been rehired." & vbCrLf & vbCrLf
        MailBodyR = MailBodyR & "New Employee #: " & glbLEE_ID & vbCrLf
        MailBodyR = MailBodyR & "Old Employee #: Unknown" & vbCrLf
        MailBodyR = MailBodyR & "Start Date: " & GetEmpData(glbLEE_ID, "ED_DOH") & vbCrLf
        MailBodyR = MailBodyR & "First Name: " & GetEmpData(glbLEE_ID, "ED_FNAME") & vbCrLf
        MailBodyR = MailBodyR & "Last Name: " & GetEmpData(glbLEE_ID, "ED_SURNAME") & vbCrLf
        
        MailBodyR = MailBodyR & "Position: " & GetJobData(clpJob.Text, "JB_DESCR") & vbCrLf
        MailBodyR = MailBodyR & lStr("Rept. Authority 1") & ": " & GetEmpData(txtReptAuthority(0).Text, "ED_SURNAME") & ", " & GetEmpData(txtReptAuthority(0).Text, "ED_FNAME") & vbCrLf
        MailBodyR = MailBodyR & lStr("Union") & ": " & GetTABLDesc("EDOR", clpCode(0).Text) & vbCrLf
        MailBodyR = MailBodyR & lStr("Location") & ": " & GetTABLDesc("EDLC", GetEmpData(glbLEE_ID, "ED_LOC")) & vbCrLf
        MailBodyR = MailBodyR & lStr("Department") & ": " & GetDeptName(clpDept.Text, "DF_NAME") & vbCrLf
        MailBodyR = MailBodyR & lStr("Division") & ": " & Get_Division_Name(clpDiv.Text, "Division_Name") & vbCrLf
        'Screen.MousePointer = DEFAULT
        Call imgEmail_ReHire
        'Screen.MousePointer = HOURGLASS
    End If
End If

ExitLine1:

fglbNew = False

'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    flgNewCancel = False
'End If

Call Display_Value

Screen.MousePointer = DEFAULT

If NewHireForms.count > 0 Then
    glbLinNewPosSal = True
End If

If (glbMulti And Action = "A" And Data1.Recordset.RecordCount > 1) Or (fraPosition.Visible And optSalary(1)) Then
    If optSalary(1) Then
        glbLinNewPosSal = True
    End If
    If glbCompSerial = "S/N - 2288W" Then
        frmESALARYMusashi.Show
    Else
    frmESALARY.Show
    End If
End If

If glbLinamar And Action = "A" And NewHireForms.count = 0 Then
    Msg = "Do you want update the employee's Payroll and Personnel information? "
    If MsgBox(Msg, 36, "info:HR") = 6 Then
     frmBasicLinamar.Show 1
    End If
End If

If glbCompSerial = "S/N - 2291W" And Action = "A" And NewHireForms.count = 0 Then
    Msg = "Do you want update the employee's demographics? "
    If MsgBox(Msg, 36, "info:HR") = 6 Then
        frmBasicSyndesis.Show 1
    End If
End If

If glbOttawaCCAC Then
    If chkCurrent(0) Then
        If IsNull(GetSHData(glbLEE_ID, "SH_PAYP", Null)) Then
            If medHours(1) = 0 And medHours(2) = 0 Then
                MsgBox "Please enter Hours/Week and Hours/Pay Period if this is an ""E""-""Exceptional Hourly"" employee."
                medHours(1).SetFocus
                Exit Sub
            Else
                If medHours(1) = 0 Then
                    MsgBox "Please enter Hours/Week if this is an ""E""-""Exceptional Hourly"" employee."
                    medHours(1).SetFocus
                    Exit Sub
                End If
                If medHours(2) = 0 Then
                    MsgBox "Please enter Hours/Pay Period if this is an ""E""-""Exceptional Hourly"" employee."
                    medHours(2).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
End If


If glbCompSerial = "S/N - 2259W" Then 'Oxford Ticket #17400
    If glbMulti Then 'check if there is no checked 'Default' position
        If Not chkActPosition.Value Then
            If Not chkDefaultBox(glbLEE_ID) Then
                'MsgBox "Default position must be assigned to this employee"    - Ticket #21256 - use label master
                MsgBox chkActPosition.Caption & " must be assigned to this employee"
            End If
        End If
    End If
End If


If gsEMAIL_ONPOSITION Then
    If Len(MailBody) > 0 Then
        Screen.MousePointer = DEFAULT
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352
            Call EmailSendingForSamuel
        Else
            Call imgEmail_Click
        End If
    End If
End If

'Ticket #25152: Macaulay Child Development Centre - Move to Performance Review screen if New Position/New Salary
'Cannot add the (fraPosition.Visible And optSalary(0)) logic because for Multi Position we are not displaying this.
'So I have commented this out as in the above rows, for any new Position added, Salary screen is called to add
'a new Salary record.
'If gSec_Inq_Performance And glbCompSerial = "S/N - 2420W" And NewHireForms.count = 0 And Action = "A" Then
'    frmEPERFORM.Show
'End If

fraPosition.Visible = False

Action = "M"

Call NextForm

If glbWFC Then 'Ticket #25927 Franks 08/26/2014 - for hrsoft missing position process
    'If NewHireForms.count = 0 Then
        If glbCandidate > 0 Then
             Call WFCHRSoftProcUpt("frmEPOSITION")
        End If
    'End If
End If

Exit Sub

Add_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_JOB_HISTORY", "Update")
Call RollBack '26July99 js
Resume Next
End Sub

Private Function chkDefaultBox(xEmpNo)
Dim rsLEmpJob As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    If glbtermopen Then
        retVal = True 'don't check for term
    Else
        SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & xEmpNo & " AND NOT (JH_CURRENT=0)"
        SQLQ = SQLQ & " AND JH_POSITION_CONTROL = 'YES' "
        rsLEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsLEmpJob.EOF Then
            retVal = True
        End If
        rsLEmpJob.Close
    End If
    chkDefaultBox = retVal

End Function
'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

'Private Sub cmdPerform_Click()
'Unload frmEPERFORM
'glbSetPer = glbSetPos
'frmEPERFORM.Show
'Unload Me
'End Sub

Private Sub cmdPerform_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String
'cmdPrint.Enabled = False
RHeading = lblEEName & "'s Position History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False
RHeading = lblEEName & "'s Position History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

'Private Sub cmdSalary_Click()
'Unload frmESALARY
'glbSetSal = glbSetPos
'frmESALARY.Show
'Unload Me
'End Sub

Private Function CurSDate()
Dim SQLQ As String
Dim HRJH_Snap As New ADODB.Recordset

CurSDate = 0    ' returns 0 if no found records

On Error GoTo JHS_Err

SQLQ = "Select HR_JOB_HISTORY.* FROM HR_JOB_HISTORY"
SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_CURRENT <>0"

oPHRS = 0
oWHRS = 0
ODHRS = 0
oJob = ""
OLambtonJob = ""
OSDATE = ""
OLeadHand = ""
OLabourCD = ""
oLABOUREDATE = Null
OReason = ""
oPayrollID = ""
oOrg = ""
oDeptNo = ""
oGLNo = ""
oStatus = ""
OLambtonJob = ""
oPayCategory = ""
oSHIFT = ""  'Ticket #12051, for ADP interface (VitalAire, WFC, ...)
oREPTAU = "" 'Ticket #12051
HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRJH_Snap.BOF And HRJH_Snap.EOF Then
    Exit Function
Else
    'Not (Town of Aurora and City of Timmins and City of Kawartha Lakes)
    If glbVadim And glbMulti And glbCompSerial <> "S/N - 2378W" And glbCompSerial <> "S/N - 2375W" And glbCompSerial <> "S/N - 2363W" Then
        If fglbNew Then
            If empPayrollID = txtPayrollID Then
                SetEmpValue (True)
            End If
            Do While Not HRJH_Snap.EOF
                If HRJH_Snap("JH_PAYROLL_ID") = txtPayrollID Then
                    CurSDate = HRJH_Snap("JH_SDATE")
                    oPHRS = HRJH_Snap("JH_PHRS")
                    ODHRS = HRJH_Snap("JH_DHRS")
                    oWHRS = HRJH_Snap("JH_WHRS")
                    oJob = HRJH_Snap("JH_JOB")
                    OSDATE = HRJH_Snap("JH_SDATE")
                    OReason = HRJH_Snap("JH_JREASON")
                    oPayrollID = HRJH_Snap("JH_PAYROLL_ID")
                    oOrg = HRJH_Snap("JH_ORG")
                    oDeptNo = HRJH_Snap("JH_DEPTNO")
                    oGLNo = HRJH_Snap("JH_GLNO")
                    oStatus = HRJH_Snap("JH_EMP")
                    oPayCategory = HRJH_Snap("JH_PAYROLL_CATEGORY")
                    If IsNull(HRJH_Snap("JH_SHIFT")) Then
                        oSHIFT = ""
                    Else
                        oSHIFT = IIf(Not IsNull(HRJH_Snap("JH_SHIFT")), HRJH_Snap("JH_SHIFT"), "")
                    End If
                    If IsNull(HRJH_Snap("JH_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRJH_Snap("JH_REPTAU")
                    If IsNull(HRJH_Snap("JH_GRID")) Then
                        OLambtonJob = oJob
                    Else
                        OLambtonJob = Left(HRJH_Snap("JH_GRID"), 1) & oJob & Mid(HRJH_Snap("JH_GRID"), 2)
                    End If
                    HRJH_Snap("JH_CURRENT") = 0
                    HRJH_Snap.Update
                End If
                HRJH_Snap.MoveNext
            Loop
    
            HRJH_Snap.Close
        Else
            CurSDate = 0
            If empPayrollID = txtPayrollID Then
                SetEmpValue
            End If
            oPHRS = Data1.Recordset("JH_PHRS")
            ODHRS = Data1.Recordset("JH_DHRS")
            oWHRS = Data1.Recordset("JH_WHRS")
            oJob = Data1.Recordset("JH_JOB")
            OSDATE = Data1.Recordset("JH_SDATE")
            OReason = Data1.Recordset("JH_JREASON")
            oPayrollID = Data1.Recordset("JH_PAYROLL_ID")
            oOrg = Data1.Recordset("JH_ORG")
            oDeptNo = Data1.Recordset("JH_DEPTNO")
            oGLNo = Data1.Recordset("JH_GLNO")
            oStatus = Data1.Recordset("JH_EMP")
            oPayCategory = Data1.Recordset("JH_PAYROLL_CATEGORY")
            If IsNull(HRJH_Snap("JH_SHIFT")) Then
                oSHIFT = ""
            Else
                oSHIFT = IIf(IsNull(Data1.Recordset("JH_SHIFT")), "", Data1.Recordset("JH_SHIFT"))
            End If
            'oREPTAU = Data1.Recordset("JH_REPTAU")
            If IsNull(Data1.Recordset("JH_REPTAU")) Then oREPTAU = "" Else oREPTAU = Data1.Recordset("JH_REPTAU")
            If IsNull(Data1.Recordset("JH_GRID")) Then
                OLambtonJob = oJob
            Else
                OLambtonJob = Left(Data1.Recordset("JH_GRID"), 1) & oJob & Mid(Data1.Recordset("JH_GRID"), 2)
            End If
        End If
    ElseIf glbMulti Then
        Do While Not HRJH_Snap.EOF
            If HRJH_Snap("JH_JOB") = clpJob.Text Then
                CurSDate = HRJH_Snap("JH_SDATE")
                oPHRS = HRJH_Snap("JH_PHRS")
                ODHRS = HRJH_Snap("JH_DHRS")
                oWHRS = HRJH_Snap("JH_WHRS")
                oJob = HRJH_Snap("JH_JOB")
                OSDATE = HRJH_Snap("JH_SDATE")
                OReason = HRJH_Snap("JH_JREASON")
                oPayrollID = HRJH_Snap("JH_PAYROLL_ID")
                oOrg = HRJH_Snap("JH_ORG")
                oDeptNo = HRJH_Snap("JH_DEPTNO")
                oGLNo = HRJH_Snap("JH_GLNO")
                oStatus = HRJH_Snap("JH_EMP")
                If IsNull(HRJH_Snap("JH_SHIFT")) Then
                    oSHIFT = ""
                Else
                    oSHIFT = HRJH_Snap("JH_SHIFT")
                End If
                If IsNull(HRJH_Snap("JH_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRJH_Snap("JH_REPTAU")
            End If
            HRJH_Snap.MoveNext
        Loop
        HRJH_Snap.Close
    Else
        CurSDate = HRJH_Snap("JH_SDATE")
        oPHRS = HRJH_Snap("JH_PHRS")
        ODHRS = HRJH_Snap("JH_DHRS")
        oWHRS = HRJH_Snap("JH_WHRS")
        oJob = HRJH_Snap("JH_JOB")
        OSDATE = HRJH_Snap("JH_SDATE")
        OReason = HRJH_Snap("JH_JREASON")
        If IsNull(HRJH_Snap("JH_SHIFT")) Then
            oSHIFT = ""
        Else
            oSHIFT = HRJH_Snap("JH_SHIFT")
        End If
        If IsNull(HRJH_Snap("JH_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRJH_Snap("JH_REPTAU")
        If glbLinamar Then
            OLeadHand = HRJH_Snap("JH_LEADHAND")
            OLabourCD = HRJH_Snap("JH_LABOURCD")
            oLABOUREDATE = HRJH_Snap("JH_LABOUREDATE")
        End If
        HRJH_Snap.Close
    End If
End If

Exit Function
JHS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job History Snap", "HR_JOB_HISTORY", "SELECT")
Call RollBack '26July99 js

End Function

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    glbFrmCaption$ = Me.Caption
    glbErrNum& = ErrorNumber
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRSTATS", "SELECT")
    Call RollBack '26July99 js
End Sub

Public Function EERetrieve()
Dim SQLQ As String
Dim X, xFld
Dim rt As New ADODB.Recordset
Dim rs As New ADODB.Recordset

EERetrieve = False

On Error GoTo EERError
    
    Screen.MousePointer = HOURGLASS
    
    
    If glbCompSerial = "S/N - 2259W" Then 'Added by Bryan 11/07/05 Ticket #8857
        If glbtermopen Then
            SQLQ = "Select ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_SECTION") = "Y" Then
            glbMulti = True
            frmMulti.Visible = True
        Else
            glbMulti = False
            frmMulti.Visible = False
        End If
        
        If glbMulti Then
            '7.9 - Commenting this because you have to use the Label Master now.
            'chkActPosition.Caption = "Default Position"
        Else
            If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #17030
                'chkActPosition.Caption = "Default Position"    - Ticket #21256 - use label master
            Else
                '7.9 - Commenting this because you have to use the Label Master now.
                'chkActPosition.Caption = "Acting Position"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If
    
    'WDGPHU - Ticket #27899
    If glbCompSerial = "S/N - 2411W" Then
        If glbtermopen Then
            SQLQ = "Select ED_ORGT1 FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_ORGT1 FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_ORGT1") = "YES" Then
            glbMulti = True
            frmMulti.Visible = True
        Else
            glbMulti = False
            frmMulti.Visible = False
        End If
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If
    
    If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab 'Ticket #14791
        locOrg = ""
        If glbtermopen Then
            SQLQ = "SELECT ED_EMPNBR, ED_ORG FROM Term_HREMP WHERE ED_EMPNBR = " & glbTERM_ID
        Else
            SQLQ = "SELECT ED_EMPNBR, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
        End If
        rt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly
        If Not IsNull(rt("ED_ORG")) Then
            locOrg = rt("ED_ORG")
        End If
        rt.Close
        Set rt = Nothing
        SQLQ = ""
    End If
    
    If glbWFC Then 'Ticket #30359 Franks 07/11/2017
        IsWFC_CONP = False
        If glbtermopen Then
            SQLQ = "SELECT ED_EMPNBR, ED_EMP FROM Term_HREMP WHERE ED_EMPNBR = " & glbTERM_ID
        Else
            SQLQ = "SELECT ED_EMPNBR, ED_EMP FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
        End If
        rt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly
        If Not IsNull(rt("ED_EMP")) Then
            If rt("ED_EMP") = "CONP" Then
                IsWFC_CONP = True
            End If
        End If
        rt.Close
        Set rt = Nothing
        SQLQ = ""
    End If
    
    If glbtermopen Then
        SQLQ = "Select Term_JOB_HISTORY.*,"
    Else
        SQLQ = "Select HR_JOB_HISTORY.*,"
    End If

    For X = 0 To 3 '2 Ticket #20052 Franks 07/14/2011
        xFld = "REPTAU" & IIf(X = 0, "", X + 1)
        If glbLinamar Then
            SQLQ = SQLQ & " CASE WHEN JH_" & xFld & " IS NOT NULL AND LEN(JH_" & xFld & ")>2 "
            SQLQ = SQLQ & " THEN RIGHT(JH_" & xFld & ",3)+'-'+"
            SQLQ = SQLQ & " LEFT(JH_" & xFld & ",LEN(JH_" & xFld & ")-3) "
            SQLQ = SQLQ & " ELSE STR(JH_" & xFld & ") END "
            SQLQ = SQLQ & " AS " & xFld & IIf(X = 3, "", ",") '2 -> 3
        Else
            If glbOracle Then
                SQLQ = SQLQ & "JH_" & xFld & " AS " & xFld & IIf(X = 3, "", ",")
            Else
                SQLQ = SQLQ & "STR(JH_" & xFld & ") AS " & xFld & IIf(X = 3, "", ",") '2 -> 3
            End If
        End If
    Next
    
    If glbLinamar Then 'Ticket #28846 Franks 08/16/2016
            SQLQ = SQLQ & ", CASE WHEN JH_SHIFT IS NOT NULL AND LEN(JH_SHIFT)>3 "
            SQLQ = SQLQ & " THEN SUBSTRING(JH_SHIFT,4,2) "
            SQLQ = SQLQ & " ELSE JH_SHIFT END "
            SQLQ = SQLQ & " AS SHIFT"
    End If
    
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY "
    
    'Ticket #21511 - County of Oxford - since they are able to switch between multi and non-multi, they are
    'seeing an issue with sort order, so this will fix it.
    If glbCompSerial = "S/N - 2259W" Then
        SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
    ElseIf glbMulti Then
        SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
    End If
    SQLQ = SQLQ & "JH_SDATE DESC, JH_ID DESC"
    
    Data1.RecordSource = SQLQ
    
    Data1.Refresh
    
'If glbGuelph And (Not glbtermopen) Then
'Dim RsTempEmp As New ADODB.Recordset
'    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & glbLEE_ID
'    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    xEFDATE = ""
'    xETDATE = ""
'    xNumVac = 0
'    If Not RsTempEmp.EOF Then
'        xNumVac = RsTempEmp("ED_VAC")
'        xEFDATE = RsTempEmp("ED_EFDATE")
'        xETDATE = RsTempEmp("ED_ETDATE")
'    End If
'    RsTempEmp.Close
'End If

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    lblHrsPayPeriod.FontBold = True
ElseIf glbCompSerial = "S/N - 2357W" And glbEmpCountry <> "CANADA" Then   'I.T. Xchange
    lblHrsPayPeriod.FontBold = False
End If

Screen.MousePointer = DEFAULT
EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Call RollBack '26July99 js

End Function


Private Sub Form_Activate()
    glbOnTop = "FRMEPOSITION"
    flgloaded = True
    Call Job_Desc
    Call SET_UP_MODE
    
    If glbWFC Then 'Ticket #22991 Franks 12/24/2012
        Call WFC_PT_PenCheck
        Call ReptsEffDatesScreen 'Ticket #29343 Franks 10/17/2016
    End If
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEPOSITION"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim rsTA As New ADODB.Recordset

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMEPOSITION"
    
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbMulti Or glbLinamar Then
    frmMulti.Visible = True
    If glbLinamar Then 'Ticket# 8293
        frmMulti.Height = 500
    End If
End If

frmJobEnd.BorderStyle = 0


Call LabelSetup

If glbCompSerial = "S/N - 2382W" Then 'Ticket #18090 Samuel
    lblReptAuth(0).FontBold = True
    'Ticket #20052 Franks 07/14/2011 - begin
    frmSamuelProfitSharing.Top = frmLinamar(0).Top
    frmSamuelProfitSharing.Left = frmLinamar(0).Left
    chkProSha.DataField = "JH_PROFIT_SHARING"
    frmSamuelProfitSharing.Visible = True
    'Ticket #20052 Franks 07/14/2011 - end
End If
If glbWFC Then 'Ticket #14927
    clpJob.TextBoxWidth = 1315 ' 1215 'Ticket #25911 Franks 11/10/2014
    'Ticket #21339 Frank WFC need these two fields show up
    'frmJobEnd.Visible = False
    frmJobEnd.Top = 2600 - 80 '3240
    lblReason.Top = 365 + 40
    clpCode(2).Top = 320 + 40
    
    lblBand.Left = frmJobEnd.Left
    clpCode(6).Left = 6750
    
    'Ticket #15396 - begin
    lblReptAuth(0).FontBold = True
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
    'Ticket #15396 - end
Else 'non WFC
    clpJob.TextBoxWidth = 1315 'Ticket #26726 Franks 06/15/2015
End If

lblBand.Visible = glbWFC
clpCode(6).Visible = glbWFC

'Ticket #20367 - Jerry wants to show the Position Group for everyone. Since we are using the WFC's
'Band field, this part not for WFC only.
If Not glbWFC Then
'Ticket #19864 Samuel
'If glbCompSerial = "S/N - 2382W" Then
    'GetJobData (clpJob.Text,"JB_GRPCD")
    lblBand.Visible = True
    clpCode(6).Visible = True
    
    If glbVadim Then
        lblBand.Top = lblBillingRate.Top
        clpCode(6).Top = medBillingRate.Top
        lblBand.Left = lblBillingRate.Left
        clpCode(6).Left = clpPayrollCategory.Left
    Else
        lblBand.Top = txtComments2.Top ' 4450
        clpCode(6).Top = txtComments2.Top '4410
        lblBand.Left = frmJobEnd.Left
        clpCode(6).Left = frmJobEnd.Left + 1500
    End If
    
    clpCode(6).Enabled = False
    lblBand.Caption = "Position Group"
    clpCode(6).TablName = "JBGC"
End If

If glbWFC Then 'Ticket #11772
    cboShift.Left = txtShift.Left
    cboShift.Visible = True
    txtShift.Visible = False
    cboShift.AddItem "NS"
    cboShift.AddItem "1"
    cboShift.AddItem "2"
    cboShift.AddItem "3"
    cboShift.AddItem "4"
    cboShift.AddItem "5"
    cboShift.AddItem "6"
    'Ticket #27298 Franks 07/10/2015 - add 7,8,9
    cboShift.AddItem "7"
    cboShift.AddItem "8"
    cboShift.AddItem "9"
    cboShift.AddItem "A"
    cboShift.AddItem "B"
    cboShift.AddItem "C"
    cboShift.AddItem "D"
    cboShift.AddItem "E"
    cboShift.AddItem "F"
    cboShift.AddItem "M"
    cboShift.AddItem "Q"
    cboShift.AddItem "R"
    cboShift.AddItem "S"
    cboShift.AddItem "T"
    cboShift.AddItem "W"
End If

If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
Else
    lblPosTitle.Top = lblGrid.Top
    clpJob.Top = clpGrid.Top
End If

'Ticket #27774 Franks 11/18/2015
If glbWFC Then
    Call DispJobCode
    Call ReptsEffDatesScreen 'Ticket #29343 Franks 10/17/2016
End If
'Ticket #23537 - Essex County Library - new fields
If glbCompSerial = "S/N - 2296W" Then
    frEssexLib.Visible = True
Else
    frEssexLib.Visible = False
End If

Call CR_Job_Snap

Screen.MousePointer = HOURGLASS

glbLinNewPosSal = False
lblSection.Caption = "Section" 'St. John's relabelled to Section - NA. So on scrolling thru one emply to another - the program is duplicating the label within itself.

'Call setCaption(lblGrid)
lblGrid.Caption = lStr("Grid Category")

Call TabOrderSetup

Screen.MousePointer = DEFAULT

If glbLinamar Then
    Call LinamarSceenSetup
    'chkActPosition.Visible = True
'ElseIf glbCompSerial = "S/N - 2391W" Then 'Ticket #26979 Franks 04/24/2015
'    Call NYCHScreenSetup
Else
    If glbMulti Then 'George on Dec 7,2005 #9928 begin
        '7.9 - This is now taken care from label master. So Multi Position clients will have to use Label Master to set this.
        'chkActPosition.Caption = "Default Position"
        
        'chkActPosition.Visible = True
    End If 'George on Dec 7,2005 #9928 end
    panControls.Height = 0
End If

If glbVadim Then
    lblHrsDay.FontBold = True
    lblPayID.FontBold = True
    lblPayrollCategory.Visible = True
    clpPayrollCategory.Visible = True
    
    'Ticket #30025 - They don't want Payment Type to show on employee's Position screen.
    lblRegion.Visible = False
    clpRegion.Visible = False
End If

If glbInsync Then
    If Not glbCompSerial = "S/N - 2411W" Then
        '2411 Wellington-Dufferin-Guelph Public Health Ticket #16625
        lblHrsDay.FontBold = True
        lblHrsWeek.FontBold = True
        lblHrsPayPeriod.FontBold = True
    End If
End If

'Ticket #24565 - Making Hours/Pay Period mandatory as it's required to compute the Salary per Pay to transfer to
'Vadim as per the new formula
'Ticket #19113 - Making Hours/Week mandatory as it's required for computing Salary per Pay and transferring to Vadim
If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
End If

'Burlington Tech
If glbCompSerial = "S/N - 2351W" Then
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
    cmdBackupPosition.Visible = True
    cmdBackupPosition.Left = 240
    panControls.Height = 540
End If

'Ticket #17786 Charton-Hobbs Inc
If glbCompSerial = "S/N - 2418W" Then
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
    lblShift.Visible = False
    txtShift.Visible = False
    cboShift.Visible = False
End If

'Granite Club
If glbCompSerial = "S/N - 2241W" Then
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
    lblTitle(9).FontBold = True
    lblTitle(2).FontBold = True
    lblTitle(0).FontBold = True
    lblUnion.FontBold = True
    lblPT.FontBold = True
    'lblSection.FontBold = True
    'lblReason.FontBold = True
    'lblPayID.FontBold = True
End If

'Hemu - CollectCorp Inc. - Ticket #14247
If glbCompSerial = "S/N - 2390W" Then
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
End If


'Hamilton CAS - Ticket #13398
If glbCompSerial = "S/N - 2257W" Then
    chkActPosition.Caption = "Red Circled"
End If

If glbLambton Then
    lblLambtonJob.Visible = True
    txtLambtonJob.Visible = True
    
    'lblLambtonJob.Left = 5580
    'txtLambtonJob.Left = 7390
    
    'lblBand.Top = 5300
    'clpCode(6).Top = 5250
    
    chkUseForBenefit.DataField = "JH_USRCHECK"
    chkUserDef(1).DataField = "" ' this is the one for linamar
End If
If glbAdv Then
    If Not glbCompSerial = "S/N - 2242W" And Not glbCompSerial = "S/N - 2390W" Then   'london ccac
        If isATIncluded(glbLEE_ID) Then
            lblShift.FontBold = True
        End If
    End If
    If glbLambton Then
        txtShift.MaxLength = 4
    End If
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    medBillingRate.DataField = "JH_BILLINGRATE"
    medBillingRate.Visible = True
    lblBillingRate.Visible = True
End If

If glbCompSerial = "S/N - 2380W" Then 'Vitalaire
    'Ticket #24976 - Label changed, and add dropdown list
    'lblShift.Caption = "Job Class"
    lblShift.Caption = "Workforce Category"
    comShift.Left = txtShift.Left
    comShift.Visible = True
    comShift.Tag = "00-Workforce Category"
    txtShift.Visible = False
    
    Call Populate_ComShift
    
    Call VitalAireJobFamilyScreen 'Ticket #26233 Franks 11/24/2014 VitalAire Canada Inc.
End If

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
    comShift.Left = txtShift.Left
    comShift.Visible = True
    comShift.Tag = "00-Shift"
    txtShift.Visible = False
    lblShift.FontBold = True
    Call Populate_ComShift
End If

If glbCompSerial = "S/N - 2225W" Then ''PowerStream Inc. (Markham Hydro) - Ticket #16560
    lblShift.Caption = "Position Budget No."
    txtShift.Tag = "00-Position Budget No."
End If

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    panControls.Height = 540
    cmdJobFiles.Visible = True
    
    If Not gSec_Inq_Job_Files_Attachment Then
        cmdJobFiles.Enabled = False
    End If
End If

'Four Villages Community Health Centre - Ticket #18221
If glbCompSerial = "S/N - 2425W" Then
    lblHrsPayPeriod.FontBold = True
End If

'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
If Not glbtermopen Then
    panControls.Height = 540
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Visible = True
    Else
        cmdReCompDAccrual.Visible = False
    End If
End If

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNION = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then      'Hemu -EXE
        If glbUNION = "EXEC" Then      'Hemu -EXE
            If Not glbCompSerial = "S/N - 2382W" Then 'Ticket #16266 Samuel
            'Samuel is using this function but they want show Position screen
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
        
    End If
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNIONTe = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then      'Hemu -EXE
        If glbUNIONTe = "EXEC" Then    'Hemu -EXE
            If Not glbCompSerial = "S/N - 2382W" Then 'Ticket #16266 Samuel
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
    End If
End If

If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls Ticket #27681 Franks 12/10/2015
    Call NiagaraFallsScreen
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Show Primary Position checkbox
    chkPrimary.Visible = True
Else
    'Hide Primary Position checkbox
    chkPrimary.Visible = False
End If

'City of Pickering - Ticket #13281
'The Walter Fedy Partnership - Ticket #14003
Dim rsTB As New ADODB.Recordset
rsTB.Open "SELECT ED_PT,ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
locFTPT = ""
If Not rsTB.EOF Then
    locFTPT = rsTB("ED_PT")
End If
rsTB.Close
If glbCompSerial = "S/N - 2217W" Or glbCompSerial = "S/N - 2386W" Then
        If locFTPT = "FT" Then
            lblHrsDay.FontBold = True
            lblHrsWeek.FontBold = True
            lblHrsPayPeriod.FontBold = True
        Else
            lblHrsDay.FontBold = False
            lblHrsWeek.FontBold = False
            lblHrsPayPeriod.FontBold = False
        End If
End If
If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #20886
    Call SamuelScreenSetup
End If

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    lblHrsPayPeriod.FontBold = True
End If

Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    frmEPOSITION.Caption = IIf(glbSetPos, "Set ", "") & "Position History - " & Left$(glbLEE_SName, 5)
    frmEPOSITION.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
Else
    frmEPOSITION.Caption = "Position History - New Employee"
    frmEPOSITION.lblEEName = " "
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)
lblEEID = glbLEE_ID

Call Job_Desc

clpGLNum.TextBoxWidth = 1500

Call INI_Controls(Me)

clpGrid.SecurityMaintainable = False

Screen.MousePointer = DEFAULT

Call setCaption(lblUnion)
Call setCaption(lblPT)
Call setCaption(lblRegion)
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(3))
Call setCaption(lblTitle(9))
Call setCaption(lblSection)
clpGrid.TABLTitle = lStr(lblGrid)

Call Display_Value

If glbCompSerial = "S/N - 2375W" Then 'Timmins
    rsTA.Open "SELECT ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
    If rsTA.EOF = False And rsTA.BOF = False Then
        If rsTA("ED_REGION") = "S" Then
            lblHrsWeek.FontBold = True
            lblHrsPayPeriod.FontBold = True
        Else
            lblHrsWeek.FontBold = False
            lblHrsPayPeriod.FontBold = False
        End If
    Else
        lblHrsWeek.FontBold = False
        lblHrsPayPeriod.FontBold = False
    End If
End If

Call InitData

Action = "M"
savWHRS = medHours(1)
savGrid = clpGrid.Text

If glbOttawaCCAC Then
    
    frmMulti.Visible = True
    
    Call ComEType
    lblSection = "Emp. Type"
    comEmpType.Visible = True
    clpCode(5).Visible = False 'section
    lblEndDATE.Visible = False 'end date
    dlpENDDATE.Visible = False 'end date
    lblReason.Visible = False 'end reason
    clpCode(2).Visible = False 'end reason
    frmOCCAC.Visible = True
    frmMulti.Height = 2300
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/22/2014 Franks
    lblPayID.Visible = False
    txtPayrollID.Visible = False
End If
End Sub

Private Sub Job_Desc()
Dim SQLQ As String
Dim X%
Dim rsJOB As New ADODB.Recordset
Dim rsJobGrade As New ADODB.Recordset
Dim rsWRK As ADODB.Recordset
On Error GoTo Jobd_Err

If Len(clpJob.Text) > 0 Then
    rsJOB.Open "SELECT * FROM HRJOB WHERE JB_CODE='" & CStr(clpJob.Text) & "'", gdbAdoIhr001, adOpenForwardOnly
    rsJobGrade.Open "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & CStr(clpJob.Text) & "' AND JB_GRID='" & clpGrid.Text & "'", gdbAdoIhr001, adOpenForwardOnly
    
    If rsJOB.EOF Then
        clpCode(6) = "": clpCode(6).Visible = False
    Else
        If glbWFC Then If IsNull(rsJOB("JB_BAND")) Then clpCode(6).Text = "" Else clpCode(6).Text = rsJOB("JB_BAND")
    End If
    
    If glbMultiGrid Then
        Set rsWRK = rsJobGrade
    Else
        Set rsWRK = rsJOB
    End If
        
    If rsWRK.EOF Then
        medFTEHCalc = Empty
    Else
        If Len(medFTEHCalc) = 5 Then
            medFTEHCalc = ""
        Else
            medFTEHCalc = rsWRK("JB_FTEHrs") & ""
        End If
        
        'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
        'For x% = 1 To 11
        'For X% = 1 To 15
        For X% = 1 To 20
            If Not IsNull(rsWRK("JB_S" & X%)) Then JobSnap_PayScale(X) = Round2DEC(rsWRK("JB_S" & X%))
        Next
        If Not IsNull(rsWRK("JB_SALCD")) Then JobSnap_Salary_Code$ = rsWRK("JB_SALCD")
        If Not IsNull(rsWRK("JB_MIDPOINT")) Then JobSnap_MidPoint! = rsWRK("JB_MIDPOINT")
        If Not IsNull(rsWRK("JB_ORG")) Then
            'Ticket #20367 - Jerry wants to show the Position Group for everyone. Since we are using the WFC's
            'Band field, this part of the code is for WFC only.
            If glbWFC Then
            'Ticket #19864 Samuel
            'If glbCompSerial <> "S/N - 2382W" Then
                clpCode(6).Visible = (rsWRK("JB_ORG") = "NONE" Or rsWRK("JB_ORG") = "EXEC") And glbWFC
            End If
        End If
    End If
Else
    'Ticket #20367 - Jerry wants to show the Position Group for everyone. Since we are using the WFC's
    'Band field, this part of the code is for WFC only.
    If glbWFC Then
    'Ticket #19864 Samuel
    'If glbCompSerial <> "S/N - 2382W" Then
        clpCode(6).Visible = False
    End If
    clpJob.ShowDescription = True
End If
If glbLambton Then
    txtLambtonJob = Left(clpGrid, 1) & clpJob & Mid(clpGrid, 2)
End If

Exit Sub

Jobd_Err:
If Err = 94 Then
    medFTEHCalc = ""
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "JOBS", "SELECT")
Call RollBack '26July99 js

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEPOSITION = Nothing
    Call NextForm
End Sub

Private Sub medEssex_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medFTEHrs_GotFocus()
    medFTEHrs.MaxLength = 7
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medFTEHrs_LostFocus()
    If Val(medFTEHCalc) <> 0 Then
        If medFTEHrs.Text = "" Then
            Exit Sub
        Else
            medFTENum.Text = medFTEHrs.Text / medFTEHCalc
        End If
    End If
End Sub

Private Sub medFTENum_GotFocus()
    medFTENum.MaxLength = 4
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medFTENum_LostFocus()
    If Len(medFTEHCalc) <> 0 Then
        If Not IsNumeric(medFTENum) Then medFTENum = 0
        medFTEHrs.Text = medFTENum.Text * medFTEHCalc
    End If
End Sub

Private Sub medHours_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Set_Current_Flag()
Dim SQLQ As String, Msg$, X
Dim dyn_HRJOBHIS As New ADODB.Recordset

On Error GoTo CurFlgErr
If glbMulti Then Exit Sub

dyn_HRJOBHIS.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dyn_HRJOBHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
gdbAdoIhr001.BeginTrans

If dyn_HRJOBHIS.RecordCount > 0 Then dyn_HRJOBHIS.MoveFirst
dyn_HRJOBHIS("JH_CURRENT") = True
'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    dyn_HRJOBHIS("JH_TRK_CRS_RENEWAL") = False
'End If
dyn_HRJOBHIS.Update

dyn_HRJOBHIS.MoveNext

While Not dyn_HRJOBHIS.EOF
    If dyn_HRJOBHIS("JH_CURRENT") <> 0 Then
        dyn_HRJOBHIS("JH_CURRENT") = False
        'Ticket #21511 - If the Current not checked then Default Position should be Off too.
        dyn_HRJOBHIS("JH_POSITION_CONTROL") = "NO"
        dyn_HRJOBHIS.Update
    End If
    dyn_HRJOBHIS.MoveNext
Wend
gdbAdoIhr001.CommitTrans

dyn_HRJOBHIS.Close

If Not glbOracle And Not glbSQL Then Pause (0.5)
Data1.Refresh

Exit Sub

CurFlgErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_JOB_HIS", "Add")
Call RollBack '26July99 js

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


If glbMulti Then
    frmMulti.Enabled = TF
    If fglbNew And glbCompSerial <> "S/N - 2378W" Then
        txtPayrollID.Enabled = True
'        chkCurrent(0).Enabled = True
    Else
        txtPayrollID.Enabled = False
''        If Not chkCurrent(0) Then
''            chkCurrent(0).Enabled = False
''        Else
''            chkCurrent(0).Enabled = TF
''        End If
    End If
Else
'    chkCurrent(0).Enabled = TF
End If
chkCurrent(0).Enabled = TF
chkTrackCrsRenewal.Enabled = TF
chkCurrent(0).Font3D = (chkCurrent(0).Enabled + 1) * 3      'jaddy 6/8/99
medFTEHrs.Enabled = TF
medFTENum.Enabled = TF
medHours(0).Enabled = TF
medHours(1).Enabled = TF
medHours(2).Enabled = TF
clpCode(1).Enabled = TF
If IsWFC_CONP Then 'Ticket #30359 Franks 07/11/2017
    clpJob.Enabled = False
Else
    clpJob.Enabled = TF
End If
txtReptAuthority(0).Enabled = TF
txtReptAuthority(1).Enabled = TF
txtReptAuthority(2).Enabled = TF
' sam add
elpReptAuthShow(0).Enabled = TF
elpReptAuthShow(1).Enabled = TF
elpReptAuthShow(2).Enabled = TF
elpReptAuthShow(3).Enabled = TF
If glbWFC Then 'Ticket #29343 Franks 10/17/2016
    dlpRptDate(1).Enabled = TF
    dlpRptDate(2).Enabled = TF
    dlpRptDate(3).Enabled = TF
    dlpRptDate(4).Enabled = TF
    comShift.Enabled = TF
    cboShift.Enabled = TF
End If
txtPosCtr.Enabled = TF
txtShift.Enabled = TF
dlpStartDate.Enabled = TF
txtComment.Enabled = TF     'Jaddy 6/4/99
txtComments2.Enabled = TF
If clpGrid.Visible Then clpGrid.Enabled = TF
''Franks Jul 11,02
'If Len(cmbBand) = 0 Then
'    cmbBand.Enabled = False
'Else
'    cmbBand.Enabled = TF
'End If
'If Len(clpCode(6)) = 0 Then
    clpCode(6).Enabled = False
'Else
'    clpCode(6).Enabled = TF
'End If
clpPayrollCategory.Enabled = TF
dlpENDDATE.Enabled = TF
clpCode(2).Enabled = TF

If glbLinamar Then
    Dim X
    For X = 0 To 4
        frmLinamar(X).Enabled = TF
        If X <> 0 And X <> 4 Then
            txtLabel(X).Visible = False
            chkUserDef(X).Enabled = Len(lblLabel(X)) <> 0 And TF
            dlpDate(X).Enabled = Len(lblLabel(X)) <> 0 And TF
        End If
    Next
'    cmdEditLable.Enabled = TF
'    chkActPosition.Enabled =tf
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Primary checkbox for Multi Position only
    chkPrimary.Enabled = TF
End If

'If glbCompSerial = "S/N - 2391W" Then 'Ticket #26979 Franks 04/24/2015
'    frmNYCH.Enabled = TF
'    clpSalDist.Enabled = TF
'End If

'If Not gSec_Inq_Salary Then cmdSalary.Enabled = False
'If Not gSec_Inq_Performance Then cmdPerform.Enabled = False
If Not gSec_Upd_Salary Then
    optSalary(1) = False
    fraPosition.Visible = False
End If

'George on Jan 26,2006 #10266
glbJob = "" 'George on Jan 24,2006 #10266
glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    If Not IsNull(Data1.Recordset("JH_JOB")) Then glbJob = Data1.Recordset("JH_JOB") 'George on Jan 19,2006 #10266
    If Not IsNull(Data1.Recordset("JH_SDATE")) Then glbSDate = Data1.Recordset("JH_SDATE") 'George on Jan 24,2006 #10266
End If
glbDocName = "Offer"
If gsAttachment_DB Then
    Call DispimgIcon(Me, "frmEPOSITION")
    If gSec_Upd_Position And Not glbtermopen Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If
'George on Jan 26,2006 #10266

'Simona - begin - Assessment Strategies-#14963
If (glbCompSerial = "S/N - 2401W") Then
    If NewHireForms.count > 0 Then
        medHours(0).Text = "7.5"
        medHours(1).Text = "37.5"
        medHours(2).Text = "75.0"
        medFTENum.Text = "1"
        medFTEHrs.Text = "1950"
    End If
End If
'Simona - end - Assessment Strategies-#14963

'Ticket #16288, use it as GP Benefit Group Code
'If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #15590
'    lblTitle(3).Enabled = False
'    clpGLNum.Enabled = False
'End If

End Sub

Private Sub medHours_LostFocus(Index As Integer)
    'City of Pickering - Ticket #13281
    If glbCompSerial = "S/N - 2217W" Then
        If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
            If IsNumeric(medHours(2)) Then  'Hours/Pay Period
                medFTEHrs = medHours(2) * 26
            End If
        End If
    End If

    'Ticket #24543 - Macaulay Child Development Centre
    If fglbNew And glbCompSerial = "S/N - 2420W" Then
        If IsNumeric(medHours(1)) Then  'Hours/Week
            'Hours/Pay Period
            medHours(2) = Round((medHours(1) * 52) / 24, 2)
        End If
    End If

End Sub

Private Sub optSalary_Click(Index As Integer, Value As Integer)
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20371 Franks 05/25/2011
    If optSalary(0).Value Then
        Call SamuelCurEDate(True)
    End If
    If optSalary(1).Value Then
        Call SamuelCurEDate(False)
    End If
End If
If glbWFC Then 'Ticket #29343 Franks 10/18/2016
    If optSalary(2).Value Then '
        Call getLastCurrPositionDat
        Call EnableNoneReptFields(False)
    Else
        Call EnableNoneReptFields(True)
    End If
End If
End Sub

Private Sub getLastCurrPositionDat()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 AND JH_EMPNBR = " & glbLEE_ID & " "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        clpJob.Text = rs("JH_JOB")
        dlpStartDate.Text = rs("JH_SDATE")
        If Not IsNull(rs("JH_REPTAU")) Then elpReptAuthShow(0).Text = rs("JH_REPTAU")
        If Not IsNull(rs("JH_REPTAU2")) Then elpReptAuthShow(1).Text = rs("JH_REPTAU2")
        If Not IsNull(rs("JH_REPTAU3")) Then elpReptAuthShow(2).Text = rs("JH_REPTAU3")
        If Not IsNull(rs("JH_REPTAU4")) Then elpReptAuthShow(3).Text = rs("JH_REPTAU4")
        If Not IsNull(rs("JH_EDATEREPT1")) Then dlpRptDate(1).Text = rs("JH_EDATEREPT1")
        If Not IsNull(rs("JH_EDATEREPT2")) Then dlpRptDate(2).Text = rs("JH_EDATEREPT2")
        If Not IsNull(rs("JH_EDATEREPT3")) Then dlpRptDate(3).Text = rs("JH_EDATEREPT3")
        If Not IsNull(rs("JH_EDATEREPT4")) Then dlpRptDate(4).Text = rs("JH_EDATEREPT4")
        clpCode(1).Text = rs("JH_JREASON")
    End If
    rs.Close
End Sub

Private Sub EnableNoneReptFields(xFlag As Boolean)

    If IsWFC_CONP Then 'Ticket #30359 Franks 07/11/2017
        clpJob.Enabled = False
    Else
        clpJob.Enabled = xFlag
    End If
    dlpStartDate.Enabled = xFlag
    medHours(0).Enabled = xFlag
    medHours(1).Enabled = xFlag
    medHours(2).Enabled = xFlag
    cboShift.Enabled = xFlag
    clpCode(1).Enabled = xFlag
    medFTENum.Enabled = xFlag
    medFTEHrs.Enabled = xFlag
    dlpENDDATE.Enabled = xFlag
    clpCode(2).Enabled = xFlag
End Sub

Private Sub scrControl_Change()
    fraDetail.Top = 60 + vbxTrueGrid.Height + panEEDESC.Height - scrControl.Value * ((panControls.Height + scrControl.Max) / scrControl.Max)
End Sub

Private Sub txtComment_GotFocus()
    Call SetPanHelp(Me.ActiveControl)       'Jaddy 6/4/99
End Sub

Private Sub txtComments2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDouDiv_Change(Index As Integer) 'Ticket #26233 Franks 11/24/2014 VitalAire Canada Inc.
lblDouDivDesc(Index).Caption = getJobFamilyDesc(txtDouDiv(Index).Text, Index)
End Sub

Private Sub txtLabCode_Change()
    lblLabCodeDesc.Caption = getLabCodeDesc(txtLabCode.Text)
End Sub
Private Function getLabCodeDesc(xCode)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = "Unassigned"
    If Not IsNull(xCode) Then
        SQLQ = "SELECT TB_NAME, TB_KEY, TB_DESC FROM HRTABL WHERE TB_NAME = 'SDLB' AND TB_KEY = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("TB_DESC")
        End If
        rsDiv.Close
    End If
    getLabCodeDesc = xRetVal
End Function

Private Sub txtLabCode_DblClick()
    Call Get_Code_Normal("SDLB", "Labour Code", "")
    If Len(glbCode) > 0 Then
        txtLabCode.Text = glbCode
    End If
End Sub

Private Sub txtLabel_Change(Index As Integer)
Dim X
If glbLinamar Then
    If Index = 1 And Len(txtLabel(1)) = 0 Then
        lblLabel(1) = "POC"
    ElseIf Index = 2 And Len(txtLabel(2)) = 0 Then
        lblLabel(2) = "LPS Program Manger"
    ElseIf Index = 3 And Len(txtLabel(3)) = 0 Then
        lblLabel(3) = "P.I. Program"
    Else
        lblLabel(Index) = txtLabel(Index)
    End If
    chkUserDef(Index).Enabled = Len(lblLabel(Index).Caption) <> 0 'And cmdOK.Enabled  'check the condition
    dlpDate(Index).Enabled = Len(lblLabel(Index).Caption) <> 0 'And cmdOK.Enabled
End If
End Sub

Private Sub txtLabel_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtLabel_LostFocus(Index As Integer)
Dim X
If glbLinamar Then
    For X = 2 To 3
        If Len(txtLabel(X)) = 0 Then chkUserDef(X) = False: dlpDate(X).Text = ""
    Next
End If

End Sub

Private Sub txtPayrollID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPayrollID_KeyPress(KeyAscii As Integer)
If glbVadim Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtPosCtr_DblClick()
 frmJobsCCAC.PosCode = Trim(clpJob.Text)
 frmJobsCCAC.Show 1
 If Not IsEmpty(frmJobsCCAC.PosNbr) Then txtPosCtr = frmJobsCCAC.PosNbr
End Sub

Private Sub txtPosCtr_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPosCtr_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtReptAuthority_Change(Index As Integer)
    elpReptAuthShow(Index) = ShowEmpnbr(txtReptAuthority(Index).Text)
End Sub

Private Sub txtShift_Change()
    If cboShift.Visible Then
        cboShift.Text = txtShift.Text
    End If
    
    If comShift.Visible Then
        If glbCompSerial = "S/N - 2380W" Then 'Vitalaire
            'Ticket #24976 - Label changed, and add dropdown list
            comShift.ListIndex = -1
            comShift.ListIndex = GetComShiftIndex(Trim(txtShift.Text))
        End If
        If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
            comShift.ListIndex = -1
            comShift.ListIndex = GetComShiftIndex(Trim(txtShift.Text))
        End If
    End If
End Sub

Private Sub txtShift_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Upd_Related_Salary()

Dim SQLQ As String, Msg As String
Dim dynHRSALHIS As New ADODB.Recordset
Dim JobCode$, PositionStartDat, JobReason$
Dim HoursPerWeek!
Dim lngJobID&

Dim X!, cX$
Dim SH_SALARY@, SH_SALCD$, SH_EDATE, SH_PAYP$, SH_NEXTDAT As Variant
Dim xSH_FISCALYEAR, xSH_SECTION, xSH_MARKETLINE, xSH_BAND, xSH_CURRENCYINDI 'WFC ONLY
Dim SHisDate, SPosDate  As Variant
Dim AnnualSalary As Double, Compa!, SalaryGrade$
Dim xPosEarly
Dim xSH_PREMIUM, xSH_TOTAL, xSH_VGROUP, xSH_VSTEP
Dim xSHID 'George added Mar 9,2006 #9965

On Error GoTo UpRel_Err

JobCode$ = clpJob.Text

If IsNumeric(Data1.Recordset("JH_ID")) Then lngJobID& = Data1.Recordset("JH_ID") Else lngJobID& = 0

If Not IsNull(dlpStartDate.Text) Then PositionStartDat = CVDate(dlpStartDate.Text)
If Not IsNull(medHours(1)) And Len(medHours(1)) > 0 Then HoursPerWeek! = medHours(1)
If Not IsNull(clpCode(1).Text) Then JobReason$ = clpCode(1).Text


SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
'Ticket #26355 - Removing the comment ORDER BY DESC because, this is Same Salary update, so pick the Salary from the last
'Current Salary. If ORDER BY DESC is removed, then if there were more than one same Salary Effective Date records, then
'older Salary is getting picked up of the Same Effective Date.
'Ticket #24096 - Change in the logic done by Ticket #24064 is causing other issues. When adding a new Salary record with
'Same Salary is updating older salary record's Start Date, i.e. the first record in this selection and other issues.
If fglbNew Then
    'Ticket #26355
    'SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " '& IIf(glbSQL Or glbOracle, "DESC", "")
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL Or glbOracle, "DESC", "")
Else
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL Or glbOracle, "DESC", "")
End If
dynHRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRSALHIS.BOF And dynHRSALHIS.EOF Then
    Msg = "No salary records found - New Employee?" & Chr(10)
    Msg = Msg & "Please review and update this Employee's" & Chr(10)
    Msg = Msg & "salary."
    MsgBox Msg
    dynHRSALHIS.Close
    Exit Sub
End If

SHisDate = CVDate(dynHRSALHIS("SH_EDATE"))
If IsNull(dynHRSALHIS("SH_SDATE")) Then 'Ticket #24074 Franks 07/17/2013
    SPosDate = Date
Else
    SPosDate = CVDate(dynHRSALHIS("SH_SDATE"))
End If

'Ticket #24096 - Since now the Start Date can be same as previous record's start date, I now have to check if Job Codes of Prv and New
'Job is same then only update the Start Date. Also when changing an existing position only.
If Not fglbNew Then
    'xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0
    xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0 And dynHRSALHIS("SH_JOB") = JobCode$
    If xPosEarly Then
        If fgtxtStartDate = SHisDate And dynHRSALHIS("SH_JOB") = JobCode$ Then
            dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
            dynHRSALHIS.Update
            Exit Sub
        End If
    End If
End If

'Ticket #24064 - Jerry said to disable this logic of adding a new salary record when Position and/or Start Date changes.
If fglbNew Then
    dynHRSALHIS("SH_CURRENT") = False
    dynHRSALHIS.Update
    xSHID = dynHRSALHIS("SH_ID")
    'George added Mar 9,2006 #9965
    'If glbCompSerial = "S/N - 2259W" Or glbGP Then
    '    Call Salary_Integration(glbLEE_ID, , False, False, xSHID)
    'End If
    'George added Mar 9,2006 #9965

    If Not IsNull(dynHRSALHIS.Fields("SH_SALARY")) Then SH_SALARY@ = dynHRSALHIS.Fields("SH_SALARY")
    If Not IsNull(dynHRSALHIS.Fields("SH_SALCD")) Then SH_SALCD$ = dynHRSALHIS.Fields("SH_SALCD")
    If Not IsNull(dynHRSALHIS.Fields("SH_PAYP")) Then SH_PAYP$ = dynHRSALHIS.Fields("SH_PAYP")
    If Not IsNull(dynHRSALHIS.Fields("SH_NEXTDAT")) Then SH_NEXTDAT = dynHRSALHIS.Fields("SH_NEXTDAT")
    If glbWFC Then
        xSH_FISCALYEAR = "": xSH_SECTION = "": xSH_MARKETLINE = "": xSH_BAND = ""
        xSH_CURRENCYINDI = ""
        If Not IsNull(dynHRSALHIS.Fields("SH_FISCALYEAR")) Then xSH_FISCALYEAR = dynHRSALHIS.Fields("SH_FISCALYEAR")
        If Not IsNull(dynHRSALHIS.Fields("SH_SECTION")) Then xSH_SECTION = dynHRSALHIS.Fields("SH_SECTION")
        If Not IsNull(dynHRSALHIS.Fields("SH_MARKETLINE")) Then xSH_MARKETLINE = dynHRSALHIS.Fields("SH_MARKETLINE")
        If Len(clpCode(6).Text) > 0 Then xSH_BAND = clpCode(6).Text
        If Not IsNull(dynHRSALHIS.Fields("SH_CURRENCYINDI")) Then xSH_CURRENCYINDI = dynHRSALHIS.Fields("SH_CURRENCYINDI")
    End If
    If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
        If Not IsNull(dynHRSALHIS.Fields("SH_PREMIUM")) Then xSH_PREMIUM = dynHRSALHIS.Fields("SH_PREMIUM")
        If Not IsNull(dynHRSALHIS.Fields("SH_TOTAL")) Then xSH_TOTAL = dynHRSALHIS.Fields("SH_TOTAL")
        If Not IsNull(dynHRSALHIS.Fields("SH_VGROUP")) Then xSH_VGROUP = dynHRSALHIS.Fields("SH_VGROUP")
        If Not IsNull(dynHRSALHIS.Fields("SH_VSTEP")) Then xSH_VSTEP = dynHRSALHIS.Fields("SH_VSTEP")
    End If

    'SET COMPA RATIO
    '================
    'Days and Months added by Bryan 30/Sep/05 Ticket#9354
    If JobSnap_Salary_Code$ = "A" Then
        If SH_SALCD$ = "H" Then
            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52
        ElseIf SH_SALCD$ = "M" Then
            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 12
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 366
            Else
                AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 265
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    ElseIf JobSnap_Salary_Code$ = "H" Then
        If SH_SALCD$ = "A" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 52
        ElseIf SH_SALCD$ = "M" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 365
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    ElseIf JobSnap_Salary_Code$ = "M" Then
        If SH_SALCD$ = "A" Then
            AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
        ElseIf SH_SALCD$ = "M" Then
            AnnualSalary = SH_SALARY@
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 366) / 12
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 365) / 12
            End If
        Else
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52 / 12
        End If
    ElseIf JobSnap_Salary_Code$ = "D" Then
        If SH_SALCD$ = "H" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) / 52
        ElseIf SH_SALCD$ = "M" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 365
            End If
        ElseIf SH_SALCD$ = "A" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 365
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    End If

    ' set COMPA RATIO
    If glbWFC Then 'Ticket #25054 Franks 02/12/2014
        Compa! = Get_WFC_COMPA_FromMaster(glbUNION, JobCode$, SH_SALARY@, dynHRSALHIS.Fields("SH_SECTION"), dynHRSALHIS.Fields("SH_MARKETLINE"), dynHRSALHIS.Fields("SH_FISCALYEAR"))
    Else
        If JobSnap_PayScale(JobSnap_MidPoint!) <> 0 And AnnualSalary <> 0 Then
            Compa! = (AnnualSalary / JobSnap_PayScale(JobSnap_MidPoint!)) * 100
        Else
            Compa! = 0
        End If
    End If
    
    If Compa! > 999.99 Then
        Compa! = 999.99
    End If

    'Determine Pay Scale individual fits into
    '==========================================
    SalaryGrade$ = "00"
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For x! = 1 To 11
    'For X! = 1 To 15
    For X! = 1 To 20
        If AnnualSalary >= JobSnap_PayScale(X) And JobSnap_PayScale(X) > 0 Then
          cX$ = CStr(X)
          If X! <= 9 Then cX$ = "0" & cX$
          SalaryGrade$ = cX$
        End If
    Next X!

    'NOW UPDATE SALARY HISTORY TABLE  - only if new record do we add record
    '================================
    If DateDiff("d", PositionStartDat, SHisDate) > 0 And glbSetPos Then GoTo SkipSal_Change

        If Not xPosEarly Then dynHRSALHIS.AddNew

        dynHRSALHIS("SH_COMPNO") = "001" 'SH_COMPNO%
        dynHRSALHIS("SH_EMPNBR") = glbLEE_ID
        dynHRSALHIS("SH_CURRENT") = True
        dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20371 Franks 05/25/2011
            dynHRSALHIS("SH_EDATE") = IIf(xPosEarly, SHisDate, CVDate(PositionStartDat))
            If optSalary(0).Value Then
                If glbMsgCustomVal = 1 Then 'keep
                    If IsDate(dlpCurSEDate.Text) Then
                        dynHRSALHIS("SH_EDATE") = CVDate(dlpCurSEDate.Text)
                    End If
                End If
            End If
        Else
            dynHRSALHIS("SH_EDATE") = IIf(xPosEarly, SHisDate, CVDate(PositionStartDat))
        End If
        dynHRSALHIS("SH_TRANSDATE") = Format(Now, "SHORT DATE")
        dynHRSALHIS("SH_SALARY") = SH_SALARY@
        dynHRSALHIS("SH_SALCD") = SH_SALCD$
        dynHRSALHIS("SH_JOB") = JobCode$
        dynHRSALHIS("SH_GRID") = clpGrid.Text
        dynHRSALHIS("SH_PAYROLL_ID") = txtPayrollID
        'lngJobID&
        dynHRSALHIS("SH_JOB_ID") = lngJobID&
        dynHRSALHIS("SH_PAYP_TABLE") = "SDPP"
        dynHRSALHIS("SH_PAYP") = SH_PAYP$
        If IsDate(SH_NEXTDAT) Then
            If CVDate(SH_NEXTDAT) > IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)) Then
                dynHRSALHIS("SH_NEXTDAT") = SH_NEXTDAT
            End If
        End If
        dynHRSALHIS("SH_WHRS") = HoursPerWeek!
        dynHRSALHIS("SH_SREAS_TABLE") = "SDRC"
        dynHRSALHIS("SH_SREAS1") = JobReason$     ' reason code
        dynHRSALHIS("SH_COMPA") = Round(Compa!, 2)
        dynHRSALHIS("SH_GRADE") = Format(SalaryGrade$, "00")
        dynHRSALHIS("SH_LDATE") = Date
        dynHRSALHIS("SH_LTIME") = Time$
        dynHRSALHIS("SH_LUSER") = glbUserID
        If glbWFC Then
            If Len(xSH_FISCALYEAR) > 0 Then
                dynHRSALHIS("SH_FISCALYEAR") = xSH_FISCALYEAR
            End If
            If Len(xSH_SECTION) > 0 Then
                dynHRSALHIS("SH_SECTION") = xSH_SECTION
            End If
            If Len(xSH_MARKETLINE) > 0 Then
                dynHRSALHIS("SH_MARKETLINE") = xSH_MARKETLINE
            End If
            If Len(xSH_BAND) > 0 Then
                dynHRSALHIS("SH_BAND") = xSH_BAND
            End If
            If Len(xSH_CURRENCYINDI) > 0 Then
                dynHRSALHIS("SH_CURRENCYINDI") = xSH_CURRENCYINDI
            End If
        End If
        If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
            If Len(xSH_PREMIUM) > 0 Then
                dynHRSALHIS("SH_PREMIUM") = xSH_PREMIUM
            End If
            If Len(xSH_TOTAL) > 0 Then
                dynHRSALHIS("SH_TOTAL") = xSH_TOTAL
            End If
            If Len(xSH_VGROUP) > 0 Then
                dynHRSALHIS("SH_VGROUP") = xSH_VGROUP
            End If
            If Len(xSH_VSTEP) > 0 Then
                dynHRSALHIS("SH_VSTEP") = xSH_VSTEP
            End If
        End If
        dynHRSALHIS.Update

SkipSal_Change:
        xSHID = dynHRSALHIS("SH_ID")

        'Ticket #27056 - Update Audit table with this new Salary record
        If Not xPosEarly Then
            Call AUDITSALY(dynHRSALHIS("SH_EMPNBR"), dynHRSALHIS("SH_SALARY"), dynHRSALHIS("SH_PAYP"), dynHRSALHIS("SH_JOB"), dynHRSALHIS("SH_GRID"), dynHRSALHIS("SH_PAYROLL_ID"), dynHRSALHIS("SH_SALCD"), dynHRSALHIS("SH_WHRS"), dynHRSALHIS("SH_EDATE"), IIf(Not IsDate(dynHRSALHIS("SH_NEXTDAT")), Null, dynHRSALHIS("SH_NEXTDAT")), dynHRSALHIS("SH_SREAS1"))
        End If

    dynHRSALHIS.Close

    Call updBenefitForSalDEPN(glbLEE_ID)

    'City of Niagara Falls - Ticket #15542
    If glbVadim And glbCompSerial = "S/N - 2276W" Then
        'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
        Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)), "", Val(SalaryGrade$), JobCode$, "A")
    End If

    'George added Mar 9,2006 #9965
    If glbCompSerial = "S/N - 2259W" Or glbGP Then 'Or (glbWFC And glbPlantCode = "GREN") Then
        Call Salary_Integration(glbLEE_ID, , False, IIf(xPosEarly, False, True), xSHID)
    End If
    'George added Mar 9,2006 #9965

End If

'Ticket #24096 - I had to add New Position only flag because it was updating existing records when new position was being added.
'Ticket #24064 - Update the Position Code and/or Position Start Date change on the related Salary and Performance
'records
If Not fglbNew Then
    Call Update_Related_SalaryPerformance_History(fgtxtjob, fgtxtStartDate)
End If


Exit Sub

UpRel_Err:
If Err = 3021 Then
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HRSAL/PERF", "INSERT")
Call RollBack '26July99 js

End Sub
'Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X As Integer
Call Display_Value
Call Job_Desc
If glbLinamar Then
    For X = 1 To 3
        If X = 1 And Len(txtLabel(1)) = 0 Then
            lblLabel(1) = "POC"
        ElseIf X = 2 And Len(txtLabel(2)) = 0 Then
            lblLabel(2) = "LPS Program Manger"
        ElseIf X = 3 And Len(txtLabel(3)) = 0 Then
            lblLabel(3) = "P.I. Program"
        Else
            lblLabel(X) = txtLabel(X)
        End If
        chkUserDef(X).Enabled = Len(lblLabel(X)) <> 0 'And cmdOK.Enabled  'should set new condition
        dlpDate(X).Enabled = chkUserDef(X).Enabled
    Next
End If
'George on Dec 7,2005 #9928 begin
If chkCurrent(0) Then
    chkActPosition.Enabled = True
Else
    chkActPosition.Enabled = False
End If
'George on Dec 7,2005 #9928 end

'glbJob = Data1.Recordset("JH_JOB") 'George on Jan 19,2006 #10266
'glbSDate = Data1.Recordset("JH_SDATE") 'George on Jan 24,2006 #10266

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function

Private Sub InitData()
If Len(clpJob.Text) > 0 Then      ' only if these change
    fgtxtjob = clpJob.Text        ' need we update sal his
End If
If Len(dlpStartDate.Text) > 0 Then
    If IsDate(dlpStartDate.Text) Then fgtxtStartDate = CVDate(dlpStartDate.Text)
End If
fgtxtDhrs = medHours(0)
End Sub

Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer
Dim rsTT As New ADODB.Recordset

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If oENDDATE <> "" Then    'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    'Macaulay Child Dev. 8.0 - Ticket #24564: Follow Up record with ENDC Code
    If glbCompSerial = "S/N - 2420W" Then
        SQLQ = SQLQ & " AND EF_FREAS = 'ENDC'"
    Else
        SQLQ = SQLQ & " AND EF_FREAS = 'RFED'"
    End If
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(oENDDATE)
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And dlpENDDATE.Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpENDDATE.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        'Macaulay Child Dev. 8.0 - Ticket #24564: Follow Up record with ENDC Code
        If glbCompSerial = "S/N - 2420W" Then
            rsTB("EF_FREAS") = "ENDC"
        Else
            rsTB("EF_FREAS") = "RFED"
        End If
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        
        'Macaulay Child Dev. 8.0 - Ticket #24564
        'Create the Follow Up Code if not existing and the Security on the Follow Up Code
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='ENDC'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "ENDC"
            rsTT("TB_DESC") = "End of Contract"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "ENDC", "End of Contract")
        
        Exit Function
    End If
    If fglbNew% = False And Edit1 = False And dlpENDDATE.Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpENDDATE.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        'Macaulay Child Dev. 8.0 - Ticket #24564: Follow Up record with ENDC Code
        If glbCompSerial = "S/N - 2420W" Then
            rsTB("EF_FREAS") = "ENDC"
        Else
            rsTB("EF_FREAS") = "RFED"
        End If
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        
        'Macaulay Child Dev. 8.0 - Ticket #24564
        'Create the Follow Up Code if not existing and the Security on the Follow Up Code
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='ENDC'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "ENDC"
            rsTT("TB_DESC") = "End of Contract"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "ENDC", "End of Contract")
        
        Exit Function
    End If
  
    If fglbNew% = False And Edit1 = True And dlpENDDATE.Text <> "" Then ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpENDDATE.Text)
            'Macaulay Child Dev. 8.0 - Ticket #24564: Follow Up record with ENDC Code
            If glbCompSerial = "S/N - 2420W" Then
                dynHRAT("EF_FREAS") = "ENDC"
            Else
                dynHRAT("EF_FREAS") = "RFED"
            End If
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If oENDDATE <> dlpENDDATE.Text Then
            Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        
        'Macaulay Child Dev. 8.0 - Ticket #24564
        'Create the Follow Up Code if not existing and the Security on the Follow Up Code
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='ENDC'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "ENDC"
            rsTT("TB_DESC") = "End of Contract"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "ENDC", "End of Contract")
        
        Exit Function
    End If
    If fglbNew% = False And Edit1 = True And dlpENDDATE.Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpENDDATE.Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
Function EMPBenefitGroup()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
EMPBenefitGroup = ""
SQLQ = "Select ED_BENEFIT_GROUP from HREMP Where ED_EMPNBR = " & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsEmp.EOF Then
    EMPBenefitGroup = rsEmp("ED_BENEFIT_GROUP") & ""
End If
End Function
Sub AddFTE(xEmpNo, xFlag)
    Dim OldFTE, NewFTE, xEFDATE, xETDATE, xNumVac
    Dim RsFTEHis As New ADODB.Recordset
    Dim xDays1, xDays2, xVacDays, xDate1, xDate2, xFDate, xTDate, xHrsDay, xHrsDayN
    Dim xVacHours, xYear, xNum As Integer, II, J
    Dim xArray(100, 2)
    Dim tNewFTE, xNumVacINS, VAC_First
    Dim RsTempEmp As New ADODB.Recordset
    Dim RsJobEmp As New ADODB.Recordset
    Dim SQLQ, xTxtJOB
    Dim FlagLoop As Boolean
    
    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xEFDATE = ""
    xETDATE = ""
    xNumVac = 0
    If Not RsTempEmp.EOF Then
        xNumVac = RsTempEmp("ED_VAC")
        xNumVacINS = RsTempEmp("ED_VAC")
        xEFDATE = RsTempEmp("ED_EFDATE")
        xETDATE = RsTempEmp("ED_ETDATE")
    End If
    RsTempEmp.Close
    
    If Len(xEFDATE) = 0 Or Len(xETDATE) = 0 Then
        Exit Sub
    End If
    
    'If xFLAG = "NEW" Then
        Call Pause(6)
    'End If
    SQLQ = "Select * from HR_JOB_HISTORY Where JH_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
    RsJobEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If RsJobEmp.EOF Then
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
    If IsDate(xEFDATE) Then
    SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
    End If
    If IsDate(xETDATE) Then
    SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
    End If
    SQLQ = SQLQ & "ORDER BY CP_FDATE DESC"
    RsFTEHis.Open SQLQ, gdbAdoSN2322, adOpenKeyset, adLockOptimistic
    If RsFTEHis.EOF And xFlag <> "NEW" Then
        Exit Sub
    End If

    If xFlag = "NEW" Then
        If xNumVac = 0 Then
            Exit Sub
        End If
        If Not RsFTEHis.EOF Then ' IF CP_VACORIGION EXIST AND CHANGE IN THE SAME YEAR
            If RsFTEHis("CP_FDATE") = xEFDATE Then
                xNumVac = RsFTEHis("CP_VACORIGION")
                GoTo MAIN_DEAL
            End If
        End If
        '' The following shows how to calculate the VAC days at the end of last year
        '' We always suppose the FTE# is 1.00 at the end of last year
        ' X is VAC days when FTE# = 1
        ' VAC_First is the first VAC days before FTE# change
        ' days1,days2, ... daysn are date range when FTE# change within this year
        ' VAC_First = X/365 * FTE#1 * days1 + X/365 * FTE#2 * days2 + ... + X/365 * FTE#n * daysn
        ' X = (VAC_First * 365)/(FTE#1 * days1 + FTE#2 * days2 + ... + FTE#n * daysn)
        VAC_First = xNumVac
        
        xDate1 = "**"
        xFDate = xEFDATE
        xTDate = xETDATE
        FlagLoop = True
        xHrsDayN = 0
        If RsJobEmp("JH_DHRS") = 0 Then
            xHrsDayN = 0
        Else
            If IsNull(RsJobEmp("JH_DHRS")) Then
                xHrsDayN = 0
            Else
                xHrsDayN = RsJobEmp("JH_DHRS")
            End If
        End If
            
        RsJobEmp.MoveNext
        II = 0
        Do While (Not RsJobEmp.EOF) And FlagLoop
            xDate1 = RsJobEmp("JH_SDATE")
            If CVDate(xDate1) > CVDate(xETDATE) Then
                GoTo Next_Rec00
            End If
            If RsJobEmp("JH_FTENUM") = 0 Then
                GoTo Next_Rec00
            End If
            If IsNull(RsJobEmp("JH_FTENUM")) Then
                GoTo Next_Rec00
            End If
            OldFTE = RsJobEmp("JH_FTENUM")
            
            If RsJobEmp("JH_DHRS") = 0 Then
                GoTo Next_Rec00
            End If
            If IsNull(RsJobEmp("JH_DHRS")) Then
                GoTo Next_Rec00
            End If
            xHrsDay = RsJobEmp("JH_DHRS")
            
            If CVDate(xDate1) < CVDate(xEFDATE) Then
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate)) * OldFTE
                FlagLoop = False
            Else
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
                xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
            End If
            
Next_Rec00:
            RsJobEmp.MoveNext
        Loop
        If IsDate(xDate1) Then
            If CVDate(xDate1) > CVDate(xEFDATE) Then
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
            End If
        End If
        
        xVacDays = 0
        For J = 1 To II
            xVacDays = xVacDays + xArray(J, 1)
        Next
        If xVacDays = 0 Then
            Exit Sub
        End If
        If xHrsDay = 0 Then
            Exit Sub
        End If
        xNumVac = Round((((VAC_First * 365) / (xVacDays)) / xHrsDayN), 0) * xHrsDayN
'        If RsFTEHis.EOF Then
'            RsFTEHis.AddNew
'        End If
'        RsFTEHis("CP_EMPNBR") = xEmpNo
'        RsFTEHis("CP_VACORIGION") = xNumVac
'        RsFTEHis("CP_VACO") = xNumVacINS
'        RsFTEHis("CP_FDATE") = xEFDATE
'        RsFTEHis("CP_TDATE") = xETDATE
'        RsFTEHis("CP_LDATE") = DATE
'        RsFTEHis("CP_LTIME") = Time$
'        RsFTEHis("CP_LUSER") = glbUSERID
'        RsFTEHis.Update
    End If
        
    If xFlag <> "NEW" Then
        If RsFTEHis.EOF Then
            xNumVac = 0
            Exit Sub
        Else
            xNumVac = RsFTEHis("CP_VACORIGION")
        End If
    End If
    
    '--- Above Got vacation days per year when FTE = 1 (xNumVac)
MAIN_DEAL:
    II = 0
    xDate1 = "**"
    xFDate = xEFDATE
    xTDate = xETDATE
    FlagLoop = True
    RsJobEmp.MoveFirst
    Do While (Not RsJobEmp.EOF) And FlagLoop
        xDate1 = RsJobEmp("JH_SDATE")
        If CVDate(xDate1) > CVDate(xETDATE) Then
            GoTo Next_Rec01
        End If
        If RsJobEmp("JH_FTENUM") = 0 Then
            GoTo Next_Rec01
        End If
        If IsNull(RsJobEmp("JH_FTENUM")) Then
            GoTo Next_Rec01
        End If
        OldFTE = RsJobEmp("JH_FTENUM")
        
        If RsJobEmp("JH_DHRS") = 0 Then
            GoTo Next_Rec01
        End If
        If IsNull(RsJobEmp("JH_DHRS")) Then
            GoTo Next_Rec01
        End If
        xHrsDay = RsJobEmp("JH_DHRS")
        
        If CVDate(xDate1) < CVDate(xEFDATE) Then
            II = II + 1
            xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate))
            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
            FlagLoop = False
        Else
            II = II + 1
            xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
            xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
            
        End If
        
Next_Rec01:
        RsJobEmp.MoveNext
    Loop
    'If IsDate(xDate1) Then
    '    If CVDate(xDate1) > CVDate(xEFDATE) Then
    '        II = II + 1
    '        xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
    '        xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
    '    End If
    'End If
    
    xVacDays = 0
    For J = 1 To II
        xVacDays = xVacDays + xArray(J, 2)
    Next
    
    If xVacDays = 0 Then
        Exit Sub
    End If
    'xVacHours = Round(xVacDays, 0) * xHrsDay
    xVacHours = Round25(xVacDays) * xHrsDay
    
    Call Pause(0.5) 'Add By Frank August 1, 2001
    
    'Dim RsTempEmp As New ADODB.Recordset
    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not RsTempEmp.EOF Then
        RsTempEmp("ED_VAC") = xVacHours
        RsTempEmp.Update
        
        If xFlag = "NEW" Then
            'If RsFTEHis.EOF Then
            RsFTEHis.AddNew
            RsFTEHis("CP_EMPNBR") = xEmpNo
            RsFTEHis("CP_VACORIGION") = xNumVac
            RsFTEHis("CP_VACO") = xNumVacINS
            RsFTEHis("CP_VACN") = xVacHours
            RsFTEHis("CP_FTENUMO") = fOldFTE
            RsFTEHis("CP_FTENUMN") = fNewFTE
            RsFTEHis("CP_FDATE") = CVDate(xEFDATE)
            RsFTEHis("CP_TDATE") = CVDate(xETDATE)
            RsFTEHis("CP_LDATE") = Date
            RsFTEHis("CP_LTIME") = Time$
            RsFTEHis("CP_LUSER") = glbUserID
            RsFTEHis.Update
            RsFTEHis.Close
            'End If
        Else
            RsFTEHis.Close
            If fOldFTE > 0 Then
            'Ticket #24677 - SQL Conversion
            'SQLQ = "DELETE * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
            SQLQ = "DELETE FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
            SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
            SQLQ = SQLQ & "AND CP_VACN = " & xNumVacINS & " "
            SQLQ = SQLQ & "AND CP_FTENUMN = " & fOldFTE & " "
            gdbAdoSN2322.Execute SQLQ
            End If
        End If
    End If
    RsTempEmp.Close
    
    
    Exit Sub

ExitLin1:
End Sub

Private Function Round25(xNumb)
Dim xInteger, xDecimal, xDecTmp
    xInteger = Int(xNumb)
    xDecimal = xNumb - xInteger
    xDecTmp = 0
    If xDecimal >= 0 And xDecimal < 0.25 Then
        xDecTmp = 0
    End If
    If xDecimal >= 0.25 And xDecimal < 0.75 Then
        xDecTmp = 0.5
    End If
    If xDecimal >= 0.75 Then
        xDecTmp = 1
    End If
    Round25 = xInteger + xDecTmp
End Function


''' Sam add July 2002 * Remove Binding Control
Public Sub Display_Value()
Dim SQLQ
Dim X, xFld

chkUseForBenefit.Visible = False 'this check box is for lambton only
chkActPosition = 0
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    
    If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
        rsDATA.CursorLocation = adUseServer
    End If
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        cmdJobFiles.Enabled = False
    End If

    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Enabled = False
    End If
Else
    If glbtermopen Then
        SQLQ = "Select Term_JOB_HISTORY.*"
    Else
        SQLQ = "Select HR_JOB_HISTORY.*"
    End If
    
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE JH_ID = " & Data1.Recordset!JH_ID
        SQLQ = SQLQ & " ORDER BY "
        'Ticket #21511 - County of Oxford - since they are able to switch between multi and non-multi, they are
        'seeing an issue with sort order, so this will fix it.
        If glbCompSerial = "S/N - 2259W" Then
            SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
        ElseIf glbMulti Then
            SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
        End If
        SQLQ = SQLQ & "JH_SDATE DESC"
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY"
        SQLQ = SQLQ & " WHERE JH_ID = " & Data1.Recordset!JH_ID
        SQLQ = SQLQ & " ORDER BY "
        
        'Ticket #21511 - County of Oxford - since they are able to switch between multi and non-multi, they are
        'seeing an issue with sort order, so this will fix it.
        If glbCompSerial = "S/N - 2259W" Then
            SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
        ElseIf glbMulti Then
            SQLQ = SQLQ & "JH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
        End If
        SQLQ = SQLQ & "JH_SDATE DESC"
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer 'Oracle version needs this
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    If glbLinamar Then
        clpJob.TransDiv = Right(glbLEE_ID, 3)
        clpCode(8).TransDiv = Right(glbLEE_ID, 3) 'Ticket #28846 Franks 07/14/2016
        'If rsDATA("JH_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1
    ElseIf glbWFC Then 'Ticket #25911 Franks 10/21/2014
        clpJob.TransDiv = glbWFCUserSecList
    ElseIf glbMulti Then 'George on Dec 7,2005 #9928 begin
        'If rsDATA("JH_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1 'George on Dec 7,2005 #9928 end
    End If
    
    If rsDATA("JH_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1
    
    Call Set_Control("R", Me, rsDATA)
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        cmdJobFiles.Enabled = True
        
        If Not gSec_Inq_Job_Files_Attachment Then
            cmdJobFiles.Enabled = False
        End If
    End If

    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Enabled = True
    End If
End If

If glbLambton Then
    If Len(clpGrid.Text) > 0 And Len(clpJob.Text) Then
        txtLambtonJob = Left(clpGrid, 1) & clpJob & Mid(clpGrid, 2)
    End If
End If
    
If chkCurrent(0) And glbLambton Then
    chkUseForBenefit.Visible = True
End If

'Ticket #20367 - Jerry wants to show the Position Group for everyone. Since we are using the WFC's
'Band field, this part not for WFC only.
If Not glbWFC Then
'Ticket #19864 Samuel
'If glbCompSerial = "S/N - 2382W" Then
    clpCode(6).Text = GetJobData(clpJob.Text, "JB_GRPCD")
End If

''If glbWFC Then 'Ticket #27774 Franks 11/18/2015
''    lblJobDesc.Caption = GetJobData(clpJob.Text, "JB_JOBCODE")
''End If

If glbCompSerial = "S/N - 2380W" Then 'Ticket #26233 Franks 11/24/2014 VitalAire Canada Inc.
    Call VitalAireJobFamilyDesc(clpJob.Text)
End If

If Not rsDATA.EOF Then Call getCodes 'Ticket #28846 Franks 07/14/2016

Call SET_UP_MODE
    
If Not glbtermopen Then
    Me.cmdModify_Click
End If
End Sub

Private Sub CR_Job_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim Job_Snap As New ADODB.Recordset

On Error GoTo Job_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT COUNT(*) AS Jobs FROM HRJOB"
Job_Snap.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Job_Snap("Jobs") = 0 Then
    Msg = "No Data in the Position Master File." & Chr(10)
    Msg = Msg & "To add Positions, go to the Menu Bar" & Chr(10)
    Msg = Msg & "and click on ''Positions''."
    MsgBox Msg
End If
Job_Snap.Close
Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Jobs", "HRJOB", "SELECT")
Call RollBack '26July99 js
 
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
'Frank Feb 9,2003, ticket 5599
'Position Security should based on Position right, not Salary right
UpdateRight = gSec_Upd_Position 'gSec_Upd_Salary
End Property

Public Property Get Addable() As Boolean
Addable = Not glbtermopen
End Property

Public Property Get Updateble() As Boolean
    Updateble = Not glbtermopen
End Property

Public Property Get Deleteble() As Boolean
Deleteble = Not glbtermopen
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEPOSITION.Caption = "Position - " & Left$(glbLEE_SName, 5)
        frmEPOSITION.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
     If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

Private Function UpdPositionCCAC()
Dim rsOC As New ADODB.Recordset
Dim rsJOBOC As New ADODB.Recordset
Dim SQLQ
If glbOttawaCCAC Then
        
    SQLQ = "SELECT JH_ID FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " AND JH_SDATE>" & Date_SQL(dlpStartDate)
    
    rsJOBOC.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsJOBOC.EOF Then Exit Function
    
    UpdPositionCCAC = False
    SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_EMPNBR =" & glbLEE_ID
    rsOC.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsOC.EOF Then
        rsOC("PC_EMPNBR") = Null
        rsOC.Update
    End If
    rsOC.Close
    
    If Len(txtPosCtr) > 0 Then
        SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_JOB='" & clpJob.Text & "'"
        SQLQ = SQLQ & " AND PC_POSITION_CONTROL='" & txtPosCtr & "'"
        rsOC.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        If rsOC.EOF Then
            rsOC.Close
            MsgBox "Invalid CCAC Position Number!"
            txtPosCtr.SetFocus
            Exit Function
        Else
            If IsNull(rsOC("PC_EMPNBR")) Then
                rsOC("PC_EMPNBR") = glbLEE_ID
                rsOC.Update
            Else
                MsgBox "The CCAC Position number has already been used."
                txtPosCtr.SetFocus
                Exit Function
            End If
            rsOC.Close
        End If
    End If
End If
UpdPositionCCAC = True
End Function

Private Sub comEmpType_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comEmpType_Click()
' 05/25/2001 Frank Modified code to add "0 - Not Applicable"
If comEmpType.ListIndex = 0 Then
    txtEmpType.Text = "0"
ElseIf comEmpType.ListIndex <> -1 Then     ' dkostka - 11/20/2001 - Added comparison to -1 to not fill in if blank.
    txtEmpType.Text = comEmpType.ListIndex
End If

End Sub

Private Sub txtEmpType_Change()
If flgloaded = False Then Exit Sub 'carmen may 00
If comEmpType.Visible = True Then
    comEmpType.ListIndex = -1
    
    If Val(txtEmpType) > 0 And Val(txtEmpType) <= 9 Then
        comEmpType.ListIndex = Val(txtEmpType)
    Else
        If txtEmpType = "0" Then
            comEmpType.ListIndex = 0
        End If
    End If
End If
End Sub

Private Sub ComEType()
comEmpType.Clear
comEmpType.AddItem "0 - Not Applicable"
comEmpType.AddItem "1 - Full Time Salary"
comEmpType.AddItem "2 - Part Time Salary"
comEmpType.AddItem "3 - Full Time Hourly"
comEmpType.AddItem "4 - Part Time Hourly"
comEmpType.AddItem "5 - Casual/Other"
comEmpType.AddItem "6 - Contract Salary"
comEmpType.AddItem "7 - Contract Hourly"    '23June99 js
comEmpType.AddItem "8 - Salary Pensioners"
comEmpType.AddItem "9 - Salary Elected officials"

End Sub

Private Sub SetDefaultValue()
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
If glbOttawaCCAC Then
    rsEmp.Open "SELECT ED_DEPTNO,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP,ED_SECTION, ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        clpDept = Format(rsEmp("ED_DEPTNO"), "@")
        clpDiv = Format(rsEmp("ED_DIV"), "@")
        clpGLNum = Format(rsEmp("ED_GLNO"), "@")
        clpCode(4) = Format(rsEmp("ED_EMP"), "@")
        clpCode(0) = Format(rsEmp("ED_ORG"), "@")
        clpPT = Format(rsEmp("ED_PT"), "@")
        clpRegion = Format(rsEmp("ED_REGION"), "@")
        If glbCompSerial = "S/N - 2332W" Then
            clpCode(5) = Format(rsEmp("ED_SECTION"), "@")
        Else
            txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@")
        End If
    End If
ElseIf glbMulti Or glbVadim Then
    rsEmp.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP,ED_PAYROLL_ID,ED_SECTION,ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
    If rsEmp.EOF Then Exit Sub
    
    clpDept = Format(rsEmp("ED_DEPTNO"), "@")
    clpDiv = Format(rsEmp("ED_DIV"), "@")
    
    If glbCompSerial = "S/N - 2259W" Then 'Oxford Ticket #15590
        clpGLNum = ""
    Else
        clpGLNum = Format(rsEmp("ED_GLNO"), "@")
    End If
    
    clpCode(4) = Format(rsEmp("ED_EMP"), "@")
    clpCode(0) = Format(rsEmp("ED_ORG"), "@")
    clpPT = Format(rsEmp("ED_PT"), "@")
    clpRegion = Format(rsEmp("ED_REGION"), "@")
    txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@")
    txtPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@")
    clpCode(5) = Format(rsEmp("ED_SECTION"), "@")
    
    If glbCompSerial = "S/N - 2362W" Or glbCompSerial = "S/N - 2379W" Then   'city of sarnia, Town of Lasalle
        clpPayrollCategory = clpDiv
    End If
    
    If glbCompSerial = "S/N - 2363W" Then ' CITY OF K LAKES
        clpPayrollCategory = rsEmp("ED_REGION") & ""
    End If
    
    'Ticket #24996 - City of Campbell River
    If glbCompSerial = "S/N - 2458W" Then
        clpPayrollCategory = clpPT
    End If
    
    empPayrollID = txtPayrollID
    rsEmp.Close
End If

 'Simona - begin - Assessment Strategies-#14963
If (glbCompSerial = "S/N - 2401W") Then
    If NewHireForms.count > 0 Then
        medHours(0).Text = "7.5"
        medHours(1).Text = "37.5"
        medHours(2).Text = "75.0"
        medFTENum.Text = "1"
        medFTEHrs.Text = "1950"
    End If
End If
'Simona - end - Assessment Strategies-#14963

'Ticket #18235 - Hours/Day/Week/Period - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    Dim xPayrollNo As String
    Dim xBranch As String
    Dim xCompany As String
    Dim xDeptno As String
    
    'Get Payroll ID and Branch #
    xPayrollNo = GetEmpData(glbLEE_ID, "ED_ADMINBY", "")
    xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
    xCompany = GetEmpData(glbLEE_ID, "ED_DIV", "")
    xDeptno = GetEmpData(glbLEE_ID, "ED_DEPTNO", "")
    Select Case xPayrollNo
        Case "2158"
            medHours(0).Text = "8"
            medHours(1).Text = "40"
            medHours(2).Text = "40"
        Case "5230"
            If xBranch <> "15" Then
                medHours(0).Text = "7.5"
                medHours(1).Text = "37.5"
                medHours(2).Text = "81.25"
            Else
                medHours(0).Text = "8"
                medHours(1).Text = "40"
                medHours(2).Text = "86.67"
            End If
        Case "5231"
            medHours(0).Text = "7.5"
            medHours(1).Text = "37.5"
            medHours(2).Text = "81.25"
        Case "5232"
            medHours(0).Text = "7.4"
            medHours(1).Text = "37"
            medHours(2).Text = "80.17"
        Case "5322"
            If xBranch <> "02" And xBranch <> "Z" Then
                medHours(0).Text = "8"
                medHours(1).Text = "40"
                medHours(2).Text = "40"
            Else
                medHours(0).Text = "12"
                medHours(1).Text = "60"
                medHours(2).Text = "60"
            End If
    End Select
    'Ticket #21652 Franks 03/20/2012
    'get hours $ rept auth from matrix
    Call SetEmpVal4Samuel(xPayrollNo, xBranch, xCompany, xDeptno)
End If

End Sub
Private Sub SetEmpVal4Samuel(xPayrollNo, xBranch, xCompany, xDeptno)
Dim rsHourRept As New ADODB.Recordset
Dim SQLQ As String
    If Len(xPayrollNo) = 0 Then Exit Sub
    If Len(xBranch) = 0 Then Exit Sub
    If Len(xCompany) = 0 Then Exit Sub
    If Len(xDeptno) = 0 Then Exit Sub
    
    'Check 4 fields
    SQLQ = "SELECT * FROM SAM_POS_ITEMS_MATRIX WHERE (1=1) "
    SQLQ = SQLQ & "AND SM_ADMINBY = '" & xPayrollNo & "' "
    SQLQ = SQLQ & "AND SM_SECTION = '" & xBranch & "' "
    SQLQ = SQLQ & "AND SM_DIV = '" & xCompany & "' "
    SQLQ = SQLQ & "AND SM_DEPTNO = '" & xDeptno & "' "
    rsHourRept.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsHourRept.EOF Then
        'check 3 fields - no deptno
        SQLQ = "SELECT * FROM SAM_POS_ITEMS_MATRIX WHERE (1=1) "
        SQLQ = SQLQ & "AND SM_ADMINBY = '" & xPayrollNo & "' "
        SQLQ = SQLQ & "AND SM_SECTION = '" & xBranch & "' "
        SQLQ = SQLQ & "AND SM_DIV = '" & xCompany & "' "
        If rsHourRept.State <> 0 Then rsHourRept.Close
        rsHourRept.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    If rsHourRept.EOF Then
        'check 2 fields - no div and deptno
        SQLQ = "SELECT * FROM SAM_POS_ITEMS_MATRIX WHERE (1=1) "
        SQLQ = SQLQ & "AND SM_ADMINBY = '" & xPayrollNo & "' "
        SQLQ = SQLQ & "AND SM_SECTION = '" & xBranch & "' "
        If rsHourRept.State <> 0 Then rsHourRept.Close
        rsHourRept.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    If Not rsHourRept.EOF Then
        If Not IsNull(rsHourRept("SM_DHRS")) Then
             medHours(0).Text = rsHourRept("SM_DHRS")
        End If
        If Not IsNull(rsHourRept("SM_WHRS")) Then
             medHours(1).Text = rsHourRept("SM_WHRS")
        End If
        If Not IsNull(rsHourRept("SM_PHRS")) Then
             medHours(2).Text = rsHourRept("SM_PHRS")
        End If
        If Not IsNull(rsHourRept("SM_REPTAU1")) Then
             txtReptAuthority(0).Text = rsHourRept("SM_REPTAU1")
        End If
        If Not IsNull(rsHourRept("SM_REPTAU2")) Then
             txtReptAuthority(1).Text = rsHourRept("SM_REPTAU2")
        End If
        If Not IsNull(rsHourRept("SM_REPTAU3")) Then
             txtReptAuthority(2).Text = rsHourRept("SM_REPTAU3")
        End If
        If Not IsNull(rsHourRept("SM_REPTAU4")) Then
             txtReptAuthority(3).Text = rsHourRept("SM_REPTAU4")
        End If
    End If
    rsHourRept.Close

End Sub

Private Sub SetEmpValue(Optional ReSetOldValue As Boolean)
Dim rsEmp As New ADODB.Recordset
Dim xUpdate As Boolean
rsEmp.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP,ED_PAYROLL_ID,ED_SECTION,ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsEmp.EOF Then Exit Sub
    If ReSetOldValue Then
        oDeptNo = Format(rsEmp("ED_DEPTNO"), "@")
'        ODIV = Format(rsEMP("ED_DIV"), "@")
        oGLNo = Format(rsEmp("ED_GLNO"), "@")
        oStatus = Format(rsEmp("ED_EMP"), "@")
        oOrg = Format(rsEmp("ED_ORG"), "@")
'        oPT = Format(rsEMP("ED_PT"), "@")
'        OEmptype = Format(rsEMP("ED_EMPTYPE"), "@")
        oPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@")
        'OSection = Format(rsEMP("ED_SECTION"), "@")
    End If
    xUpdate = False
    If clpDept <> Format(rsEmp("ED_DEPTNO"), "@") Then xUpdate = True
    If clpDiv = Format(rsEmp("ED_DIV"), "@") Then xUpdate = True
    If clpGLNum = Format(rsEmp("ED_GLNO"), "@") Then xUpdate = True
    If clpCode(4) = Format(rsEmp("ED_EMP"), "@") Then xUpdate = True
    If clpCode(0) = Format(rsEmp("ED_ORG"), "@") Then xUpdate = True
    If clpPT = Format(rsEmp("ED_PT"), "@") Then xUpdate = True
    If clpRegion = Format(rsEmp("ED_REGION"), "@") Then xUpdate = True
    If txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@") Then xUpdate = True
    If txtPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@") Then xUpdate = True
    If clpCode(5) = Format(rsEmp("ED_SECTION"), "@") Then xUpdate = True
    If xUpdate = True Then
        rsEmp("ED_DEPTNO") = clpDept
        rsEmp("ED_DIV") = clpDiv
        rsEmp("ED_GLNO") = clpGLNum
        rsEmp("ED_EMP") = clpCode(4)
        rsEmp("ED_ORG") = clpCode(0)
        rsEmp("ED_PT") = clpPT
        rsEmp("ED_REGION") = clpRegion
        rsEmp("ED_EMPTYPE") = txtEmpType
        rsEmp("ED_PAYROLL_ID") = txtPayrollID
        rsEmp("ED_SECTION") = clpCode(5)
    End If
    
    
End Sub

Private Sub UpdOttawaCCAC()
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
rsEmp.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DEPTEDATE,ED_DIVEDATE,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsEmp.EOF Then Exit Sub
If (clpDept <> rsEmp("ED_DEPTNO") And Len(clpDept) > 0) Or (clpGLNum <> rsEmp("ED_GLNO") And Len(clpGLNum) > 0) Then
    rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN": rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA("AU_NEWEMP") = "N"
    If (clpDept <> rsEmp("ED_DEPTNO") And Len(clpDept) > 0) Then
        rsTA("AU_OLDDEPT") = rsEmp("ED_DEPTNO")
        rsTA("AU_DEPTNO") = clpDept
    End If
    If rsEmp("ED_GLNO") <> clpGLNum Then
        If clpGLNum.Text <> "" Then
            rsTA("AU_DEPT_GL") = clpGLNum.Text
        Else
            rsTA("AU_DEPT_GL") = Null
        End If
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA.Update
    rsTA.Close
End If
If Len(clpDept) > 0 Then
    rsEmp("ED_DEPTNO") = clpDept
    If clpDept <> rsEmp("ED_DEPTNO") Then
        rsEmp("ED_DEPTEDATE") = dlpStartDate
    End If
End If

If Len(clpDiv) > 0 Then
    rsEmp("ED_DIV") = clpDiv
    If clpDiv <> rsEmp("ED_DIV") Then
        rsEmp("ED_DIVEDATE") = dlpStartDate
    End If
End If

If Len(clpGLNum) > 0 Then rsEmp("ED_GLNO") = clpGLNum
If Len(clpCode(4)) > 0 Then rsEmp("ED_EMP") = clpCode(4)
If Len(clpCode(0)) > 0 Then rsEmp("ED_ORG") = clpCode(0)
If Len(clpPT) > 0 Then rsEmp("ED_PT") = clpPT
If Len(clpRegion) > 0 Then rsEmp("ED_REGION") = clpRegion
If Len(txtEmpType) > 0 Then rsEmp("ED_EMPTYPE") = txtEmpType

rsEmp("ED_EMPNBR") = glbLEE_ID
rsEmp.Update
   
End Sub

Private Sub setGridList()
    
Dim rsGrid As New ADODB.Recordset
Dim xGridList As String
Dim SaveGrid As String
If Not glbMultiGrid Then Exit Sub
SaveGrid = clpGrid
clpGrid = ""
If Len(clpJob.Text) > 0 Then
    rsGrid.Open "SELECT JB_ID,JB_GRID FROM HRJOB_GRADE WHERE JB_CODE='" & CStr(clpJob.Text) & "'", gdbAdoIhr001, adOpenForwardOnly
    xGridList = ""
    Do Until rsGrid.EOF
        xGridList = xGridList & "," & rsGrid("JB_GRID")
        rsGrid.MoveNext
    Loop
    If xGridList <> "" Then xGridList = Mid(xGridList, 2)
    clpGrid.seleEMPCode = xGridList
    rsGrid.Close
Else
    clpGrid.seleEMPCode = "NONE-GRID"
End If
clpGrid = SaveGrid
End Sub

Private Function chkBenefitPayID()
Dim rsTemp As New ADODB.Recordset
Dim xID
If fglbNew Then
    xID = 0
Else
    xID = Data1.Recordset!JH_ID
End If
chkBenefitPayID = False
rsTemp.Open "SELECT JH_ID FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_USRCHECK<>0 AND JH_EMPNBR=" & glbLEE_ID & " AND JH_ID<>" & xID & " AND JH_PAYROLL_ID<>'" & txtPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
If Not rsTemp.EOF Then
    chkBenefitPayID = True
End If
End Function

Private Function GetDoh(xEmpNo)
Dim rs As New ADODB.Recordset
Dim SQLQ
    GetDoh = ""
    SQLQ = "SELECT ED_EMPNBR,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        GetDoh = rs("ED_DOH")
    End If
    rs.Close
    
End Function

Private Sub updateOMERS()
    'added by Bryan for Timmins 22/sep/05 Ticket#9368
    Dim retVal As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT ED_EMPNBR, ED_DOB, ED_DEPTNO, ED_OMERS, ED_DOH, ED_NORMALR FROM HREMP "
    strSQL = strSQL & "WHERE ED_EMPNBR = " & glbLEE_ID
    rs.Open strSQL, gdbAdoIhr001, adOpenForwardOnly, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If chkCurrent(0).Value = True Then
            Select Case clpPayrollCategory.Text
            Case "001", "002", "003", "004", "005"
                If rs("ED_DEPTNO") = "1510" Or rs("ED_DEPTNO") = "1600" Then
                   'fire and police departments retire at 60
                   If Not IsNull(rs("ED_DOB")) Then
                        retVal = DateAdd("yyyy", 60, rs("ED_DOB"))
                    End If
                Else
                    'the rest retire at 65
                    If Not IsNull(rs("ED_DOB")) Then
                        retVal = DateAdd("yyyy", 65, rs("ED_DOB"))
                    End If
                End If
                rs("ED_NORMALR") = retVal
                rs.Update
            End Select
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function fgetSection(xID) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As String
    
    If glbtermopen Then
        strSQL = "SELECT ED_SECTION FROM TERM_HREMP WHERE TERM_SEQ =" & xID
        rs.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
    Else
        strSQL = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & xID
        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    End If
    
    If rs.EOF = False Then
        If Not IsNull(rs("ED_SECTION")) Then
            retVal = rs("ED_SECTION")
        Else
            retVal = ""
        End If
    Else
        retVal = ""
    End If
    rs.Close
    Set rs = Nothing
    
    fgetSection = retVal

End Function

Private Function Get_DayHours_for_Job(xJob)
    Dim rsHRJOB As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT JB_CODE, JB_DHRS FROM HRJOB WHERE JB_CODE = '" & xJob & "'"
    rsHRJOB.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRJOB.EOF Then
        If Not IsNull(rsHRJOB("JB_DHRS")) And rsHRJOB("JB_DHRS") <> "" Then
            Get_DayHours_for_Job = rsHRJOB("JB_DHRS")
        Else
            Get_DayHours_for_Job = ""
        End If
    End If
    rsHRJOB.Close
    
End Function

Public Sub Update_Employee_Job_Training_List(xJob, xPosType, Optional xStartEndDate, Optional xEndDate)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim xDWMY, xorgPosType, xorgJob As String
    Dim SQLQ  As String
    Dim flgUnqForPos, flgNoPrvRnwl, flgNoCurRnwl, flgCrsTakenBefore, flgProcCalled As Boolean
    Dim xPrvEndDate
    Dim xComments As String
    
    On Error GoTo Employee_Job_Training_Err

    'Note: If tracking is for the Previous Job then any courses for this job which does not have
    'Previous Renewal defined should be removed for this position or
    'If tracking is for Current Job then any courses for this job which does not have
    'Current Renewal defined should be removed for this position
    
    'if this procedure is called from another procedure and not an event
    If IsMissing(xStartEndDate) Then
        flgProcCalled = False
        xorgPosType = xPosType
        xorgJob = xJob
        xStartEndDate = ""
        xEndDate = ""
    Else
        flgProcCalled = True
    End If
    
    'Get the list of Required Courses for the Job
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJob & "'"
    
    'Ticket #25609 - Training Plan by Department
    'Only courses matching employee's Department if the Course has Department Code assigned
    SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(glbLEE_ID, "ED_DEPTNO") & "'))"
    
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Ticket #25609 - Training Plan by Department
            'Check if the Course has Department assigned. If so then check if the Department of the Course matches
            'employee's Department
            'If Not IsNull(rsReqCourse("PC_DEPTNO")) And rsReqCourse("PC_DEPTNO") <> "" Then
            '    If rsReqCourse("PC_DEPTNO") <> GetEmpData(glbLEE_ID, "ED_DEPTNO") Then
            '        'Skip this course as Employee does not belong to the department this Course is setup for
            '        GoTo Next_Required_Course
            '    End If
            'End If
        
            'Check if this required course is Unique for each Position.
            'If so, then it will have to be added in the Training List even
            'though the Course code already exists for this employee for other positions
            flgUnqForPos = False
            flgNoPrvRnwl = False
            flgNoCurRnwl = False
            SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY, ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsCourseMst.EOF Then
                flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
                
                If flgUnqForPos = False And rsCourseMst("ES_RENEW_FOLLOWUP") = 99 And rsCourseMst("ES_FLWUP_PRD_DWMY") = "Y" Then
                    'Skip this course
                    GoTo Next_Required_Course
                End If
            Else
                'Course not defined in the Course Code Master - skip this course
                GoTo Next_Required_Course
            End If
            'rsCourseMst.Close
            'Set rsCourseMst = Nothing
            
            'Follow Up Effective Date Period is mandatory. Check if it exists otherwise the logic below will give an error.
            If IsNull(rsReqCourse("PC_RENEW_FOLLOWUP")) Or rsReqCourse("PC_RENEW_FOLLOWUP") = "" Then
                'Follow Up Effective Date renewal Period missing
                GoTo Next_Required_Course
            End If
                        
            'Add the Required Courses in the Training List
            'if it does not already exists for this employee or Unique for each Position
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            If flgUnqForPos <> 0 Then
                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                'If xPosType = "Previous" And chkTrackCrsRenewal And chkCurrent(0) Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                'Else
                '    If chkTrackCrsRenewal And chkCurrent(0) Then
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                '    Else
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & Left(xPosType, 1) & "'"
                '    End If
                'End If
            End If
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHRTrain.EOF Then
                'TRAINING RECORD DOES NOT EXISTS - ADD NEW ONE
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                If flgUnqForPos <> 0 Then
                    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                End If
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                End If
                
                
                'May be Training List accidently deleted or messed up
                'if the Course is Previous and procedure not called from another procdure then
                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
                If flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'The first record gets it
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                                If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
                                    'Previous Position requires this course
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                    xEndDate = rsEmpJob("TW_ENDDATE")
                                    xPosType = "Previous"
                                End If
                            Else
                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                        xPosType = "Current"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                Else
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            Else
                                xPosType = "Temporary"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            End If
                        End If
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                Else
                    'if Current then do not do anything as Current record takes precedence
                End If
                
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'If xPosType = "Current" Or (xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Then
                
                'If Course was taken and it's Position is Current then
                'make sure Current Renewal Period is there otherwise do not add the course
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'Changed
                If (flgCrsTakenBefore = True And (xPosType = "Current" Or xPosType = "Temporary") And (Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR"))) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0) Or _
                    (flgCrsTakenBefore = False And (xPosType = "Current" Or xPosType = "Temporary")) Or (flgCrsTakenBefore = True And xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Or _
                    (flgCrsTakenBefore = False And xPosType = "Previous") Then
                    
                    'Add Training Record
                    rsHRTrain.AddNew
                    rsHRTrain("TR_COMPNO") = "001"
                    rsHRTrain("TR_EMPNBR") = glbLEE_ID
                    rsHRTrain("TR_CRSCODE") = rsReqCourse("PC_CRSCODE")
                    
                    If flgCrsTakenBefore = False Then
                        If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                            'Current Course Renewal found
                            xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                            Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                                
                            'For courses not taken and are now Previous, the renewal date is based
                            'on Follow Up Renewal Period and not Previous Renewal Period - above
                            'ElseIf xPosType = "Previous" Then
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpStartDate.Text))
                            End If
                        Else    'No Current Course Renewal Period
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            'ElseIf xPosType = "Previous" Then
                            '    'For courses not taken and are now Previous, the renewal date is based
                            '    'on Follow Up Renewal Period and not Previous Renewal Period.
                            '    'If there is no current renewal then it's based on End Date only and
                            '    'Prev Renewal Period - for courses taken.
                            '    'Compute Renewal with Position End Date because there is no Current Renewal Period defined
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
                            End If
                        End If
                    Else    'Course Has Been Taken Before
                        'Course has been taken before, compute Renewal Date based on Course Taken Date
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        ElseIf xPosType = "Previous" Then
                            Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsContEdu("ES_DATCOMP")))
                            Else
                                If IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate) Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(xEndDate))
                                End If
                            End If
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        End If
                        
                        'Update Continuing Education with new Renewal Date
                        rsContEdu("ES_JOB") = xJob
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                    
                    rsHRTrain("TR_JOB") = xJob
                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                        rsHRTrain("TR_SDATE") = dlpStartDate.Text
                    Else
                        rsHRTrain("TR_SDATE") = xStartEndDate
                    End If
                    If xPosType = "Current" Then
                        rsHRTrain("TR_POS_TYPE") = "C"
                    ElseIf xPosType = "Temporary" Then
                        rsHRTrain("TR_POS_TYPE") = "T"
                    ElseIf xPosType = "Previous" Then
                        rsHRTrain("TR_POS_TYPE") = "P"
                    End If
                    'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain("TR_LUSER") = glbUserID
                    
                    'Add a Follow Up record for this Training course
'                    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
'                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    rsFollowUp.AddNew
'                    rsFollowUp("EF_COMPNO") = "001"
'                    rsFollowUp("EF_EMPNBR") = glbLEE_ID
'                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                    rsFollowUp("EF_FREAS_TABL") = "FURE"
'                    'Ticket #24257 - Do not update Admin By for them only
'                    If glbCompSerial <> "S/N - 2262W" Then
'                        rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
'                        rsFollowUp("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
'                    End If
'                    rsFollowUp("EF_FREAS") = "EDUC"
'                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                    rsFollowUp("EF_LDATE") = Date
'                    rsFollowUp("EF_LTIME") = Time$
'                    rsFollowUp("EF_LUSER") = glbUserID
'                    rsFollowUp.Update
                    
                    'Ticket #24300
                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                    
                    rsHRTrain.Update
                    
                    'rsFollowUp.Close
                    'Set rsFollowUp = Nothing
                
                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
                rsContEdu.Close
                Set rsContEdu = Nothing
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
            Else
                'TRAINING RECORD FOUND
                
                'May be Training List accidently deleted or messed up
                'if the Course is Previous and procedure not called from another procdure then
                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
                If flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'The first record gets it
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                                If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
                                    'Previous Position requires this course
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                    xEndDate = rsEmpJob("TW_ENDDATE")
                                    xPosType = "Previous"
                                End If
                            Else
                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                        xPosType = "Current"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                Else
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            Else
                                xPosType = "Temporary"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            End If
                        End If
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                Else
                    'if Current then do not do anything as Current record takes precedence
                End If
                
                
                
                'Training record for this course already exists so update the Renewal Date
                'Check which Type of Position is assigned to this course
                If rsHRTrain("TR_POS_TYPE") = "C" Then
                    'Currently the course is holding Primary Current Position Code
                    'Check which type of position requires this course
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                
                                    'Clear the Follow Up Id on the other current position rec in the Temp Position table
                                    'Search HR_JOB_HISTORY table for this Position record
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up Id on the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        Else
                            'Do not do anything because Training List has most recent Position Start Date
                        End If
                    ElseIf xPosType = "Previous" Then
                        'CURRENT - Previous
                        'Current Job becoming Previous
                        'Previous Primary Position is being tracked but Current Primary Position has this course
                        'Check if the Position in HR_TRAIN is same this Position
                        If (rsHRTrain("TR_JOB") <> xJob) Or (rsHRTrain("TR_JOB") = xJob And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(dlpStartDate.Text) And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate))) Then
                            'Do not do anything because Current takes the priority
                        Else
                            'Renewal Date based on last Course Taken date if present
                            'otherwise Follow Up Effective Date Period
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Change the renewal dates if Previous renewal is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                            
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Previous
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                    'Currently the Temporary Current Position is holding this course
                    'Check which type of position requires this course now
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        'If CVDate(rsHRTrain("TR_SDATE")) <= CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                                        
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
                                    'Search HR_TEMP_WORK table for this Position record
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'TEMPORARY - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                        
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        'Else
                            'Do not do anything because Training List has most recent Position Start Date
                        'End If
                    ElseIf xPosType = "Previous" Then
                        'TEMPORARY - Previous
                        'Do not do anything because Training List record is of the Current
                        'Temporary/Cross Training Position
                    
'                        'Previous Primary Position is being tracked but Temp. Current Position is holding this course
'                        'Check if the Position in HR_TRAIN is same this Position
'                        If rsHRTrain("TR_JOB") <> xJob Then
'                            'Do not do anything because Current takes the  priority
'                        Else
'                            'Change the renewal dates if Previous renewal is defined
'                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
'                                'No Previous Renewal Period defined so delete this job from this previous position.
'                                'It should not be in the training list for any previous job
'                                flgNoPrvRnwl = True
'                            Else
'                                'Renewal Date based on last Course Taken date if present
'                                'otherwise Follow Up Effective Date Period
'                                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
'                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
'                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
'                                End If
'                            End If
'                            If flgNoPrvRnwl = False Then
'                                'Previous Renewal period available
'                                rsHRTrain("TR_JOB") = xJob
'                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
'                                ''If Renewal date is greater than today's date then clear the Course Taken Date
'                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
'                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
'                                'End If
'                                rsHRTrain("TR_LDATE") = Date
'                                rsHRTrain("TR_LUSER") = glbUserID
'                                rsHRTrain("TR_LTIME") = Time$
'                                rsHRTrain.Update
'
'                                'Update Follow Up record - Effective Date
'                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
'                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsFollowUp.EOF Then
'                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                                    rsFollowUp("EF_LDATE") = Date
'                                    rsFollowUp("EF_LUSER") = glbUserID
'                                    rsFollowUp("EF_LTIME") = Time$
'                                    rsFollowUp.Update
'                                End If
'                                rsFollowUp.Close
'                                Set rsFollowUp = Nothing
'
'                                'Update Position record with Follow Up ID
'                                'if the course code is TRAIN
'                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                    'Search HR_JOB_HISTORY table for this Position record
'                                    'and update with Follow Up Id
'                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("TW_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'
'                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
'                                    'Search HR_TEMP_WORK table for this Position record
'                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("TW_FOLLOWUP_ID") = Null
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'                                End If
'                            Else
'                                'No Previous renewal found for this course
'
'                                'Clear the Renewal date for this course and for this employee from
'                                'Continuing Education screen
'                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
'                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
'                                SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsContEdu.EOF Then
'                                    rsContEdu("ES_RENEW") = Null
'                                    rsContEdu("ES_LDATE") = Date
'                                    rsContEdu("ES_LUSER") = glbUserID
'                                    rsContEdu("ES_LTIME") = Time$
'                                    rsContEdu.Update
'
'                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
'                                        'Since the course was completed - mark the Follow Up as
'                                        'Completed instead of deleting it.
'                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP"))
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'                                    Else
'                                        'Delete the Follow Up record for this training record
'                                        'as no Course completion record found
'                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'
'                                        'Clear the Follow Up ID in the Position record
'                                        'if the course code is TRAIN
'                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                            'Search HR_TEMP_WORK table for this Position record
'                                            'and clear with Follow Up Id
'                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                            If Not rsTJob.EOF Then
'                                                rsTJob("TW_FOLLOWUP_ID") = Null
'                                                rsTJob.Update
'                                            End If
'                                            rsTJob.Close
'                                            Set rsTJob = Nothing
'                                        End If
'                                    End If
'                                Else
'                                    'Delete the Follow Up record for this training record
'                                    'as no Course completion record found
'                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                    gdbAdoIhr001.Execute SQLQ
'
'                                    'Clear the Follow Up ID in the Position record
'                                    'if the course code is TRAIN
'                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                        'Search HR_TEMP_WORK table for this Position record
'                                        'and clear with Follow Up Id
'                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                        If Not rsTJob.EOF Then
'                                            rsTJob("TW_FOLLOWUP_ID") = Null
'                                            rsTJob.Update
'                                        End If
'                                        rsTJob.Close
'                                        Set rsTJob = Nothing
'                                    End If
'                                End If
'                                rsContEdu.Close
'                                Set rsContEdu = Nothing
'
'                                'Delete this Training List record as the course is not required by other positions
'                                SQLQ = "DELETE FROM HR_TRAIN"
'                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
'                                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                gdbAdoIhr001.Execute SQLQ
'                            End If
'                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                    'Previous Primary or Temporary position is holding this course
                    If xPosType = "Current" Then
                        'This course is required by new Current Primary Position so recalculate the
                        'Renewal Dates
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                                
                                'Check if the Course was taken before ever
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                'SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'Course Taken Before
                                    flgNoCurRnwl = True
                                Else
                                    'Course not taken before
                                    flgNoCurRnwl = False
                                    xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                    Else
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            Else
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            End If
                        Else
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            End If
                        End If
                        If flgNoCurRnwl = False Then
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            rsHRTrain("TR_JOB") = xJob
                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
                            End If
                            rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                            ''If Renewal date is greater than today's date then clear the Course Taken Date
                            'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                            '    rsHRTrain("TR_COURSE_TAKEN") = Null
                            'End If
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                                
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Current renewal found for this course - Correct logic - confirmed with email -March 09, 2009 1:18 PM
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                'Ticket #26211 Franks 11/04/2014
                                'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                    ElseIf xPosType = "Previous" Then
                        'PREVIOUS - Previous
                        'Track for the most recent previous position requiring this course
                        'These courses are for new Previous Primary Position so recalculate the
                        'Renewal Dates
                        xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                        If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                        'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate), dlpStartDate.Text, xStartEndDate)) Then
                        If (dlpENDDATE.Text = "") And (IsNull(xEndDate) Or xEndDate = "" Or IsMissing(xEndDate)) Then
                        Else
                        If CVDate(xPrvEndDate) < CVDate(IIf(IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate), dlpENDDATE.Text, xEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Previous Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                    
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                            
                                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                                                
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        
                        Else
                            'Do not do anything because Training List has most recent Position Start Date
                        End If
                        End If
                    End If
                ElseIf IsNull(rsHRTrain("TR_POS_TYPE")) Or rsHRTrain("TR_POS_TYPE") = "" Then
                    'Check if the course was taken before. If taken then use the normal Training List logic based
                    'on the renewal date if the course should continue to exist or not
                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                        'COURSE NEVER TAKEN BEFORE
                        'It's an independent course and never taken before, update with this Position's information
                        'even though there is no renewal period for the type of position this is
                        rsHRTrain("TR_JOB") = xJob
                        If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                            rsHRTrain("TR_SDATE") = dlpStartDate.Text
                        Else
                            rsHRTrain("TR_SDATE") = xStartEndDate
                        End If
                        If xPosType = "Current" Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf xPosType = "Temporary" Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        ElseIf xPosType = "Previous" Then
                            rsHRTrain("TR_POS_TYPE") = "P"
                        End If
    
                        'Do not overwrite the Renewal Date entered for this independent course
                        'rsHRTrain("TR_RENEW")) =
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain("TR_LTIME") = Time$
                        
                        'If follow up id is null then find the id
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                            SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                        
                        'Ticket #24300
                        'rsHRTrain.Update
                        
                        'Ticket #24300
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            'Add a Follow Up record for this Training course
                            'Ticket #24300
                            'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                            rsHRTrain.Update
                        Else
                            rsHRTrain.Update
                        
                            'Update Follow Up record - Comments with Position
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                'No change to renewal date
                                'rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                    
                        'Update Position record with Follow Up ID
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Clear the Follow Up from the Previous Job in Primary/Temp Position
                            'Search HR_JOB_HISTORY table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                            'Search HR_TEMP_WORK table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("TW_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    Else
                        'COURSE TAKEN BEFORE
                        'Which kind of Position is this
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        ElseIf xPosType = "Previous" Then
                            'Check if Previous Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                'No Previous Renewal Period defined so delete this job from this previous position.
                                'It should not be in the training list for any previous job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        End If
                        
                        If flgNoCurRnwl = False Then
                            'Renewal Period Found - updated existing records
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'No Job - Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Renewal Period available
                            rsHRTrain("TR_JOB") = xJob
                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
                            End If
                            If xPosType = "Current" Then
                                rsHRTrain("TR_POS_TYPE") = "C"
                            ElseIf xPosType = "Temporary" Then
                                rsHRTrain("TR_POS_TYPE") = "T"
                            ElseIf xPosType = "Previous" Then
                                rsHRTrain("TR_POS_TYPE") = "P"
                            End If
                            
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                            
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Renewal Period found for this course
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                'Ticket #26211 Franks 11/04/2014
                                'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (TR_JOB = '' OR TR_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                        
                    End If
                End If
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
                
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
Next_Required_Course:
            rsCourseMst.Close
            Set rsCourseMst = Nothing

            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

Exit Sub

Employee_Job_Training_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Employee_Job_Training_List", "HR_JOB_HISTORY", "Update_Emp_Job_Training_List")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Update_Emp_Job_Training_List")
End If
Call RollBack '26July99 js
End Sub

Private Sub Track_Courses_Renewal_Update(Optional xDelete, Optional xPosType, Optional xOldPosCode)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim flgRequired As Boolean
    Dim PosCode As String
    
    On Error GoTo Track_Courses_Renewal_Err
    
    If IsMissing(xOldPosCode) Then
        PosCode = clpJob.Text
    Else
        PosCode = xOldPosCode
    End If
    
    If chkTrackCrsRenewal And IsMissing(xDelete) Then  'Previous Position course being tracked
        'Turn-ON the tracking
        Call Update_Employee_Job_Training_List(PosCode, "Previous")
        
    Else
        'Turn-OFF the tracking
        '-------------------------------------------------------------------------------------------------
        'remove Renewal Date from the Continuing Education screen
        'retrieve Unique for each Position courses first
        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_DATCOMP,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND ES_JOB = '" & PosCode & "'"
        SQLQ = SQLQ & " AND ES_CRSCODE IN (SELECT TR_CRSCODE FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
        SQLQ = SQLQ & " ORDER BY ES_RENEW"
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            rsContEdu.MoveFirst
            
            Do While Not rsContEdu.EOF
                SQLQ = "SELECT * FROM HR_TRAIN"
                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "'"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHRTrain.EOF Then
                    If (rsHRTrain("TR_RENEW") = rsContEdu("ES_RENEW")) And (IsNull(rsHRTrain("TR_COURSE_TAKEN")) Or (rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP"))) Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    
                        If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            If Not IsMissing(xPosType) Then
                                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                            End If
                            'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete "Unique for each Position" courses from Follow Up records
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            If Not IsMissing(xPosType) Then
                                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                            End If
                            'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        
                            'Clear the Follow Up ID in the Position record
                            'if the course code is TRAIN
                            If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    End If
                Else
                    'Delete "Unique for each Position" courses from Follow Up records
                    'as no Course completion record found
                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    If Not IsMissing(xPosType) Then
                        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                    End If
                    'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                    gdbAdoIhr001.Execute SQLQ
                
                    'Clear the Follow Up ID in the Primary Position record
                    'if the course code is TRAIN
                    If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
                rsHRTrain.Close
                Set rsHRTrain = Nothing
                
                rsContEdu.MoveNext
            Loop
        Else
            'Delete "Unique for each Position" courses from Follow Up records
            'as no Course completion record found
            SQLQ = "DELETE FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
            If Not IsMissing(xPosType) Then
                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
            End If
            SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
            gdbAdoIhr001.Execute SQLQ
        End If
        rsContEdu.Close
        Set rsContEdu = Nothing
        
        'from the Training List
        SQLQ = "DELETE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
        If Not IsMissing(xPosType) Then
            SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
        gdbAdoIhr001.Execute SQLQ
        '-------------------------------------------------------------------------------------------------
        
        'Rest of this position's required courses which are not 'unique for each position'
        'Retrieve the Required Courses for this position - Non Unqiue for each Position courses
        SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & PosCode & "'"
        SQLQ = SQLQ & " AND PC_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
        
        'Ticket #25609 - Training Plan by Department
        'Only courses matching employee's Department if the Course has Department Code assigned
        SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(glbLEE_ID, "ED_DEPTNO") & "'))"
        
        rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsReqCourse.EOF Then
            
            rsReqCourse.MoveFirst
            flgRequired = False
            
            Do While Not rsReqCourse.EOF
                'Initialise if this course is required by any other position
                flgRequired = False
                
                'Check if the each required courses for this position is also required by other positions
                'Select all current positions in HR_JOB_HISTORY and HR_TEMP_WORK, and
                'Previous Positions with Tracking ON - for this employee
                SQLQ = "SELECT JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                'and not the position currently selected
                SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmpJob.EOF Then
                    rsEmpJob.MoveFirst
                    
                    Do While Not rsEmpJob.EOF
                        'Check in the Required Courses table if the retrieved required course is required by other retrieved position
                        SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & rsEmpJob("TW_JOB") & "'"
                        SQLQ = SQLQ & " AND PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                        rsPosCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsPosCourse.EOF Then
                            'Check if this course has Current and/or Previous Renewal Period
                            If rsEmpJob("TW_CURRENT") And (IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0) Then
                                'Current Position - no Current Renewal Period
                                'Check if another position required this course
                                GoTo Next_Position
                            ElseIf rsEmpJob("TW_TRK_CRS_RENEWAL") And (IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0) Then
                                'Previous Position - no Previous Renewal Period
                                'Check if another position required this course
                                GoTo Next_Position
                            End If
                            
                            'Required by another position. Do not delete this Course
                            flgRequired = False 'Changed
                                              
                            rsPosCourse.Close
                            Set rsPosCourse = Nothing
                                              
                            'Move to the next Course
                            GoTo Next_RequiredCourse
                        End If
Next_Position:
                        rsPosCourse.Close
                        Set rsPosCourse = Nothing
                        
                        rsEmpJob.MoveNext
                    Loop
                End If
Next_RequiredCourse:
                rsEmpJob.Close
                Set rsEmpJob = Nothing

                If flgRequired Then
                    'Call procedure to update Renewal Date and Position Code, Follow Up effective date
                    'Do not do anything now. At the end of this loop go through each of the
                    'courses and update the Renewal Dates and Position Codes and create the follow up records.
                Else
                    'This course is not required by any other position this employee is holding
                    'or the Current and/or Previous Renewal Period is missing which means
                    'any other position Current/Previous requiring this course will not be
                    'able to renew it without the appropriate renewal period.
                    
                    'Clear the Renewal date for this course and for this employee from
                    'Continuing Education screen
                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND ES_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    'Ticket #26211 Franks 10/29/2014
                    'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    'Ticket #26211 Franks 10/29/2014
                    'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsContEdu.EOF Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete the Follow Up record for this training record
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            'Ticket #26211 Franks 11/04/2014
                            'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        
                            'Clear the Follow Up ID in the Temp/Cross Training Position record
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    Else
                        'Delete the Follow Up record for this training record
                        'as no Course completion record found
                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                        'Ticket #26211 Franks 11/04/2014
                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                        gdbAdoIhr001.Execute SQLQ
                    
                        'Clear the Follow Up Id in the Temp/Cross Training Position record
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    End If
                    rsContEdu.Close
                    Set rsContEdu = Nothing
                    
                    'Delete this Training List record as the course is not required by other positions
                    SQLQ = "DELETE FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    gdbAdoIhr001.Execute SQLQ
                End If
                rsReqCourse.MoveNext
            Loop
        End If
        rsReqCourse.Close
        Set rsReqCourse = Nothing
        
        'Call procedure to update Renewal Dates and Position Codes and create/update the follow up records.
        'For the remaining required courses for this position which are required by other positions.
        Call Update_Remaining_Tracked_Courses(PosCode)
    End If
    
Exit Sub

Track_Courses_Renewal_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Track_Courses_Renewal_Update", "HR_JOB_HISTORY", "Courses_Renewal_Update")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Courses_Renewal_Update")
End If
Call RollBack '26July99 js
End Sub

Private Sub Update_Remaining_Tracked_Courses(xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseCode As New ADODB.Recordset
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDWMY As String
    Dim xRenewalDt
    Dim xComments As String
    
    On Error GoTo Remaining_Tracked_Courses_Err
    
    'Retrieve the Required Courses for this position - Non Unqiue for each Position courses
    'SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & clpJob.Text & "'"
    'SQLQ = SQLQ & " AND PC_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
    'rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'If Not rsReqCourse.EOF Then
    '    rsReqCourse.MoveFirst
        
    '    Do While Not rsReqCourse.EOF
            'Select all current positions in HR_JOB_HISTORY and HR_TEMP_WORK, and
            'Previous Positions with Tracking ON - for this employee
            'The records will be ordered by Current, Temporary, Previous tracked
            SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            'and not the position currently selected
            SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
            SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            'SQLQ = SQLQ & " ORDER BY POS_TYPE ASC,TW_CURRENT DESC"
            SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
            rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJobs.EOF Then
                rsEmpJobs.MoveFirst
                
                Do While Not rsEmpJobs.EOF
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        'Changed
                        'Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Current", rsEmpJobs("TW_SDATE"))
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), IIf(rsEmpJobs("POS_TYPE") = "CURRENT", "Current", "Temporary"), rsEmpJobs("TW_SDATE"))
                    Else
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Previous", rsEmpJobs("TW_SDATE"), rsEmpJobs("TW_ENDDATE"))
                    End If
                    GoTo next_EmpJob
                    
                    'Find out which position requires this course and update the training list accordingly.
                    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & rsEmpJobs("TW_JOB") & "'"
                    SQLQ = SQLQ & " AND PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    rsPosCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsPosCourse.EOF Then
                        'If Primary - CURRENT or TEMPORARY - Current
                        'Get Current Renewal Period from Course Code Master to calculate Renewal Date
                        If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                            SQLQ = "SELECT ES_CRSCODE,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY,ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            rsCourseCode.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsCourseCode.EOF Then
                                'Update Training List record with new renewal date, position, position start date, type of position
                                'Update Follow Up record and Continuing Education record as well
                                SQLQ = "SELECT * FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsHRTrain.EOF Then
                                    'Keep the original Renewal Date for record retrieval from
                                    'Continuing Education screen
                                    xRenewalDt = rsHRTrain("TR_RENEW")
                                    
                                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                        Select Case rsCourseCode("ES_FLWUP_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_FOLLOWUP"), CVDate(rsEmpJobs("TW_SDATE")))
                                    Else
                                        Select Case rsCourseCode("ES_CUR_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    End If
                                    rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                                    rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = Left(rsEmpJobs("POS_TYPE"), 1)
                                    ''If Renewal date is greater than today's date then clear the Course Taken Date
                                    'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                    '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                    'End If
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain("TR_LTIME") = Time$
                                    
                                    'Ticket #24300
                                    'rsHRTrain.Update
                                    
                                    'If follow up id is null then find the id
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                        SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    rsReqCourse.Close
                                    Set rsReqCourse = Nothing
                                    
                                    'Ticket #24300
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Add a Follow Up record for this Training course
                                        'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsCourseCode("ES_CRSCODE"), rsEmpJobs("TW_JOB"))
                                        
                                        rsHRTrain.Update
                                    Else
                                        rsHRTrain.Update
                                    
                                        'Update Follow Up record - Effective Date
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                            rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update the Continuing Education record for this course and this employee
                                    'with Renewal Date and Job Code
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND ES_JOB = '" & clpJob.Text & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(xRenewalDt)
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                    
                                    'Update Temp/Cross Training Position record with Follow Up ID
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsHRTrain.Close
                                Set rsHRTrain = Nothing
                            End If
                            rsCourseCode.Close
                            Set rsCourseCode = Nothing
                        
                        ElseIf (Not rsEmpJobs("TW_CURRENT")) And rsEmpJobs("TW_TRK_CRS_RENEWAL") Then
                            'If PREVIOUS
                            'Get Previous Renewal Period from Course Code Master to calculate the Renewal Date
                            SQLQ = "SELECT ES_CRSCODE,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY,ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY,ES_RENEW_FOLLOWUP,ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            rsCourseCode.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsCourseCode.EOF Then
                                'Update Training List record with new renewal date, position, position start date, type of position
                                'Update Follow Up record and Continuing Education record as well
                                SQLQ = "SELECT * FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsHRTrain.EOF Then
                                    'Keep the original Renewal Date for record retrieval from
                                    'Continuing Education screen
                                    xRenewalDt = rsHRTrain("TR_RENEW")
                                    
                                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                        Select Case rsCourseCode("ES_FLWUP_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_FOLLOWUP"), CVDate(rsEmpJobs("TW_SDATE")))
                                    Else
                                        Select Case rsCourseCode("ES_PRV_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        If Not IsNull(rsCourseCode("ES_RENEW_CRS_CUR")) And rsCourseCode("ES_RENEW_CRS_CUR") <> 0 Then
                                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                        Else
                                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_PRV"), CVDate(rsEmpJobs("TW_ENDDATE")))
                                        End If
                                    End If
                                    rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                                    rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "P"
                                    'If Renewal date is greater than today's date then clear the Course Taken Date
                                    'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                    '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                    'End If
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain("TR_LTIME") = Time$
                                    
                                    'Ticket #24300
                                    'rsHRTrain.Update
                                    
                                    'If follow up id is null then find the id
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                        SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    rsReqCourse.Close
                                    Set rsReqCourse = Nothing
                                    
                                    'Ticket #24300
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Add a Follow Up record for this Training course
                                        'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), rsEmpJobs("TW_JOB"))
                                        rsHRTrain.Update
                                    Else
                                        rsHRTrain.Update
                                    
                                        'Update Follow Up record - Effective Date
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                            rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update the Continuing Education record for this course and this employee
                                    'with Renewal Date and Job Code
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND ES_JOB = '" & clpJob.Text & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(xRenewalDt)
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                
                                    'Update Temp/Cross Training Position record with Follow Up ID
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                Else
                                    'Hemu - Testing
                                    'Data1.Recordset.Find "JH_ID=" & rsEmpJobs("TW_ID")
                                    'Call Display_Value
                                    'Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Previous")
                                End If
                                rsHRTrain.Close
                                Set rsHRTrain = Nothing
                            End If
                            rsCourseCode.Close
                            Set rsCourseCode = Nothing
                        End If
                    End If
                    rsPosCourse.Close
                    Set rsPosCourse = Nothing
next_EmpJob:
                    rsEmpJobs.MoveNext
                Loop
            End If
            rsEmpJobs.Close
            Set rsEmpJobs = Nothing
            
    '        rsReqCourse.MoveNext
    '    Loop
    'End If
    'rsReqCourse.Close
    'Set rsReqCourse = Nothing

Exit Sub

Remaining_Tracked_Courses_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Remaining_Tracked_Courses", "HR_JOB_HISTORY", "Remaining_Tracked_Courses")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Remaining_Tracked_Courses")
End If
Call RollBack '26July99 js
End Sub

Private Sub Update_Position_Start_Date_in_Training_List(oldSDate, newSDate)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDWMY As String
    Dim xComments As String
    
    On Error GoTo Position_Start_Date_in_Training_Err
    
    'Retrieve Training List records which match this employee, job and original start date
    SQLQ = "SELECT * FROM HR_TRAIN "
    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
    SQLQ = SQLQ & " AND TR_SDATE = " & Date_SQL(oldSDate)
    SQLQ = SQLQ & " AND TR_POS_TYPE <> 'T'"
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        rsHRTrain.MoveFirst
        
        'Training records found
        Do While Not rsHRTrain.EOF
            rsHRTrain("TR_SDATE") = newSDate    'New Date
            rsHRTrain("TR_LDATE") = Date
            rsHRTrain("TR_LUSER") = glbUserID
            rsHRTrain("TR_LTIME") = Time$
                        
            'Recompute Renewal Dates for courses with no Completion Date
            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                'Retrieve the Renewal Periods
                SQLQ = "SELECT * FROM HR_JOB_COURSE"
                SQLQ = SQLQ & " WHERE PC_JOB = '" & clpJob.Text & "'"
                SQLQ = SQLQ & " AND PC_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "'"
                rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                If Not rsReqCourse.EOF Then
                    'Ticket #25609 - Training Plan by Department
                    'Check if the Course has Department assigned. If so then only get the rest of the info from the
                    'Required Courses table if the Employee also belongs to the same Department
                    If Not IsNull(rsReqCourse("PC_DEPTNO")) And rsReqCourse("PC_DEPTNO") <> "" Then
                        If rsReqCourse("PC_DEPTNO") = GetEmpData(glbLEE_ID, "ED_DEPTNO") Then
                            'Employee belongs to the same Department
                            'Course Not Taken - Renewal Date based on Follow Up Period
                            Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(newSDate))
                        End If
                    Else
                        'Course Not Taken - Renewal Date based on Follow Up Period
                        Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(newSDate))
                    End If
                End If
                                
                'If follow up id is null then find the id
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                rsReqCourse.Close
                Set rsReqCourse = Nothing
                                
                'Ticket #24300
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Add a Follow Up record for this Training course
                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), rsHRTrain("TR_JOB"))
                Else
                    'Update Follow Up record - Effective Date
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                
            End If
            
            rsHRTrain.Update
            
            rsHRTrain.MoveNext
        Loop
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing

Exit Sub

Position_Start_Date_in_Training_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Position_Start_Date_in_Training_List", "HR_JOB_HISTORY", "Position_Start_in_Training_List")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Position_Start_in_Training_List")
End If
Call RollBack '26July99 js
End Sub

Private Function Get_Position_End_Date(xJob, xStartDate)
    Dim rsEmpJob As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_Position_End_Date = ""
    
    SQLQ = "SELECT JH_ID, JH_EMPNBR, JH_SDATE, JH_ENDDATE FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " AND JH_SDATE = " & Date_SQL(xStartDate)
    SQLQ = SQLQ & " AND JH_TRK_CRS_RENEWAL<>0"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        Get_Position_End_Date = rsEmpJob("JH_ENDDATE")
    Else
        rsEmpJob.Close
        Set rsEmpJob = Nothing
        SQLQ = "SELECT TW_ID, TW_EMPNBR, TW_SDATE, TW_ENDDATE FROM HR_TEMP_WORK"
        SQLQ = SQLQ & " WHERE TW_JOB = '" & xJob & "'"
        SQLQ = SQLQ & " AND TW_SDATE = " & Date_SQL(xStartDate)
        SQLQ = SQLQ & " AND TW_TRK_CRS_RENEWAL<>0"
        rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpJob.EOF Then
            Get_Position_End_Date = rsEmpJob("TW_ENDDATE")
        Else
            Get_Position_End_Date = ""
        End If
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
End Function

Private Sub Update_Courses_to_Next_Appropriate_Position()
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String

    On Error GoTo Next_Appropriate_Position_Err

    'Retrieve the Required Courses for this position - Non Unqiue for each Position courses
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & clpJob.Text & "'"
        
    'Ticket #25609 - Training Plan by Department
    'Only courses matching employee's Department if the Course has Department Code assigned
    SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(glbLEE_ID, "ED_DEPTNO") & "'))"
    
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            'and not the position currently selected
            SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
            SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
            SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
            SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
            rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJobs.EOF Then
                rsEmpJobs.MoveFirst
                
                Do While Not rsEmpJobs.EOF
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), IIf(rsEmpJobs("POS_TYPE") = "CURRENT", "Current", "Temporary"), rsEmpJobs("TW_SDATE"))
                    Else
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Previous", rsEmpJobs("TW_SDATE"), rsEmpJobs("TW_ENDDATE"))
                    End If
                    
                    rsEmpJobs.MoveNext
                Loop
            End If
            rsEmpJobs.Close
            Set rsEmpJobs = Nothing
        
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

Exit Sub

Next_Appropriate_Position_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Courses_to_Next_Appropriate_Position", "HR_JOB_HISTORY", "Course_Next_Position")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Course_Next_Position")
End If
Call RollBack '26July99 js
End Sub

Private Function PrvSDate(xCurrent)
Dim SQLQ As String
Dim HRJH_Snap As New ADODB.Recordset

PrvSDate = 0    ' returns 0 if no found records

On Error GoTo PrvSDate_Err

SQLQ = "Select JH_EMPNBR, JH_SDATE FROM HR_JOB_HISTORY"
SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_EMPNBR = " & glbLEE_ID & " "
If xCurrent Then
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_CURRENT =0"
End If
SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRJH_Snap.BOF And HRJH_Snap.EOF Then
    Exit Function
Else
    PrvSDate = HRJH_Snap("JH_SDATE")
End If
HRJH_Snap.Close
Set HRJH_Snap = Nothing

Exit Function
PrvSDate_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Previous Job History", "HR_JOB_HISTORY", "SELECT")
Call RollBack '26July99 js
End Function

Private Sub TabOrderSetup()
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        clpJob.TabIndex = 1
        clpGrid.TabIndex = 2
        dlpStartDate.TabIndex = 3
        elpReptAuthShow(0).TabIndex = 4
        medHours(0).TabIndex = 5
        medHours(1).TabIndex = 6
        medHours(2).TabIndex = 7
        clpCode(1).TabIndex = 8
        
        'Ticket #20371 Franks 05/25/2011
        Call SamuelNewPosScreenSetup
    End If
End Sub

Private Sub SamuelNewPosScreenSetup() 'Ticket #20371 Franks 05/25/2011
    fraPosition.Height = 975 'before: 585
    fraPosition.Width = 3945 'before: 3225
    Call SamuelCurEDate(True)
End Sub
Private Sub SamuelCurEDate(xFlag) 'Ticket #20371 Franks 05/25/2011
    lblCurSDate.Visible = xFlag
    dlpCurSEDate.Visible = xFlag
End Sub
Private Sub CheckReptAuth() 'Ticket #20885 Franks 11/10/2011 for Samuel
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
    xFlag1 = False
    'Position Change
    If Len(oJob) > 0 And Len(clpJob.Text) > 0 Then
        If oJob <> clpJob.Text Then
            'check if this employee is a Reporting Authority
            If IsReportAuth(glbLEE_ID) Then
                xFlag1 = True
            End If
        End If
    End If

    If xFlag1 Then
        xMsg = "This employee has been assigned as a Reporting Authority on other employee files."
        xMsg = xMsg & " Will this new Position affect the Reporting Authority structures?"
        frmMsgYesNoUn.lblMsg.Caption = xMsg
        frmMsgYesNoUn.lblMsg.Alignment = 0
        frmMsgYesNoUn.Show 1
        If glbMsgCustomVal = 1 Or glbMsgCustomVal = 3 Then
            'create a report to show the employee list
            Call CreateEmpList4ReportAuth(glbLEE_ID)
            'show the report - begin
            Me.vbxCrystal.Reset
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
            End If
            'Ticket #21669 Franks 03/01/2012
            'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & lblEEName & "'"
            xMsg = Replace(lblEEName, "'", "''")
            Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
            Me.vbxCrystal.Connect = RptODBC_SQL
            Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & lblEEName
            Me.vbxCrystal.Destination = 0
            Me.vbxCrystal.Action = 1
            Me.vbxCrystal.Reset
            'show the report - end
        End If
    End If
    
End Sub

Private Sub SamuelScreenSetup()
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String

If Left(locFTPT, 2) = "FT" And (glbUNION = "NONE" Or glbUNION = "EXEC") Then
    lblReptAuth(3).FontBold = True
    lblTitle(1).FontBold = True
Else
    lblReptAuth(3).FontBold = False
    lblTitle(1).FontBold = False
End If

glbEmpDiv = ""
glbEmpAdminBy = ""
glbEmpSection = ""
glbEmpRegion = ""
SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_DIV, ED_SECTION, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ADMINBY")) Then glbEmpAdminBy = "" Else glbEmpAdminBy = rsEmpee("ED_ADMINBY")
    If IsNull(rsEmpee("ED_SECTION")) Then glbEmpSection = "" Else glbEmpSection = rsEmpee("ED_SECTION")
    If IsNull(rsEmpee("ED_REGION")) Then glbEmpRegion = "" Else glbEmpRegion = rsEmpee("ED_REGION")
End If
rsEmpee.Close

End Sub

Private Sub SAMUEL_Trans(xEmpNo)
Dim xLDate
    If Len(oJob) > 0 And fglbNew Then
        If Not oJob = clpJob.Text Then
            xLDate = dlpStartDate.Text  'Date
            Call SamuelAuditAdd(xEmpNo, "M", "Position History", "New Position", oJob, clpJob.Text, xLDate)
        End If
    End If
End Sub

Private Sub LabelSetup()
    chkActPosition.Caption = lStr(chkActPosition.Caption)
    'Call setCaption(lblShift)
    lblShift.Caption = lStr("PShift")
    If lblShift.Caption = "PShift" Then lblShift.Caption = "Shift"
    Call setCaption(lblComment)
    Call setCaption(lblComment2)
    
    'Ticket #21462 Franks 02/09/2012
    lblReptAuth(0).Caption = lStr("Rept. Authority 1")
    lblReptAuth(1).Caption = lStr("Rept. Authority 2")
    lblReptAuth(2).Caption = lStr("Rept. Authority 3")
    lblReptAuth(3).Caption = lStr("Rept. Authority 4")
    vbxTrueGrid.Columns(5).Caption = lStr("Rept. Authority 1")
    vbxTrueGrid.Columns(6).Caption = lStr("Rept. Authority 2")
    vbxTrueGrid.Columns(7).Caption = lStr("Rept. Authority 3")
    vbxTrueGrid.Columns(8).Caption = lStr("Rept. Authority 4")
    
    'Ticket #23537 and Release 8.0
    lblHrsDay.Caption = lStr("Hours/Day")
    lblHrsWeek.Caption = lStr("Hours/Week")
    lblHrsPayPeriod.Caption = lStr("Hours/Pay Period")
    lblFTEHrs.Caption = lStr("FTE Hours/Year")
    vbxTrueGrid.Columns(9).Caption = lStr("Hours/Day")
    vbxTrueGrid.Columns(10).Caption = lStr("Hours/Week")
    vbxTrueGrid.Columns(11).Caption = lStr("Hours/Pay Period")
    vbxTrueGrid.Columns(14 + 1).Caption = lStr("FTE Hours/Year")
    
    'Release 8.0 - Ticket #2268: Add Payroll ID to Label Master
    lblPayID.Caption = lStr("Payroll ID")
    
    If glbLinamar Then 'Ticket #28846 Franks 08/16/2016
        vbxTrueGrid.Columns(12).Visible = False
    Else
        vbxTrueGrid.Columns(13).Visible = False
    End If
End Sub

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    If gsEMAIL_ONPOSITION Then
        If Not UserEmailExist Then
            Exit Sub
        End If

        xToEmail = GetComPreferEmail("EMAIL_ONPOSITION", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONPOSITION")
        End If
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352, do not cc it to employee
            'Else
            '    frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            'End If
            'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
            'Ticket #18578
            'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice - " & lblEEName.Caption
            'Ticket #18755
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = GetTABLDesc("EDSE", xBranch)
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Position Change Notice - " & xBranch & lblEEName.Caption
            frmSendEmail.txtSubject.Text = xEmailSubject
        
            frmSendEmail.txtBody.Text = MailBody
            'frmSendEmail.Show 1
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
            Else
                Unload frmSendEmail
            End If
            MDIMain.panHelp(0).Caption = ""
            MDIMain.panHelp(0).FloodType = 1
        End If

    End If
    Exit Sub

Email_Err:
    'If Err.Number = 364 Then
    '    Exit Sub
    'End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "EmailSendingForSamuel")
    'Resume Next
    Exit Sub

End Sub

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    If gsEMAIL_ONPOSITION Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONSALARY")
            
        'Ticket #20317 - Send email to More Emails list as well.
        xToEmail = GetComPreferEmail("EMAIL_ONPOSITION", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONPOSITION")
        End If
        
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
        
        'Samuel Ticket #18352, do not cc it to employee
        'Ticket #18856 - Friesens Corporation - do not cc it to the employee
        If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2279W" Then
        Else
            frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
        End If
        'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Position Change Notice - " & lblEEName.Caption
        frmSendEmail.txtBody.Text = MailBody
        frmSendEmail.Show 1

    End If
    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Private Sub imgEmail_NewHire()
Dim xEmail, xToEmail

On Error GoTo NewHireEmail_Err
        
    If Not UserEmailExist Then
        Exit Sub
    End If
    'xEmail = GetCurEmpEmail
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", lblEENum.Caption) 'glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
        End If
    Else
        'Ticket #20317 - More Emails for everyone
        xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
        End If
    End If
    If Len(xToEmail) > 0 Then
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONNEWHIRE")
        'frmSendEmail.txtCC.Text = xEmail
        'frmSendEmail.txtSubject.Text = "info:HR Employee New Hire Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Employee New Hire Notice - " & lblEEName.Caption
        frmSendEmail.txtBody.Text = MailBodyN
        If glbWFC Then
            'MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            '   otherwise refuse to terminate the employee.
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
                AbortTerm = False
            Else
                Unload frmSendEmail
                AbortTerm = True
            End If
            MDIMain.panHelp(0).Caption = ""
            'MDIMain.panHelp(0).FloodType = 1
        Else
            frmSendEmail.Show 1
        End If
    End If

Exit Sub

NewHireEmail_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail NewHire", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Public Sub imgEmail_ReHire()
Dim xEmail
Dim xToEmail As String
Dim EID&
On Error GoTo ReHireEmail_Err
        If Not UserEmailExist Then
            Exit Sub
        End If

        'Ticket #20317 - More Emails for everyone
        xToEmail = GetComPreferEmail("EMAIL_ONREHIRE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONREHIRE")
        End If
            
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONREHIRE")
        'frmSendEmail.txtCC.Text = xEmail
        'frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice - " & lblEEName.Caption
        frmSendEmail.txtBody.Text = MailBodyR
        frmSendEmail.Show 1

    Exit Sub

ReHireEmail_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail ReHire", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Private Function CheckDuplCurrent(xEmpNo, xJobCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & " AND JH_JOB = '" & xJobCode & "' "
    SQLQ = SQLQ & " AND JH_CURRENT <>0 " 'current
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retVal = True
    End If
    rsTemp.Close
    CheckDuplCurrent = retVal
End Function

''Private Sub WFCPosSkillsUpd(xEmpNo, xJobCode, xStartDate)
''Dim rsMain As New ADODB.Recordset
''Dim rsTemp As New ADODB.Recordset
''Dim rsEmpSki As New ADODB.Recordset
''Dim SQLQ As String
''    SQLQ = "SELECT * FROM HRJOBSKL WHERE JS_CODE = '" & xJobCode & "' "
''    SQLQ = SQLQ & "AND JS_EXPFACT = 0 "
''    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
''    Do While Not rsMain.EOF
''        'open another record setup with all records for this skill, insert it to employee skill table
''        SQLQ = "SELECT * FROM HRJOBSKL WHERE JS_CODE = '" & xJobCode & "' "
''        SQLQ = SQLQ & "AND JS_SKILL = '" & rsMain("JS_SKILL") & "' "
''        If rsTemp.State <> 0 Then rsTemp.Close
''        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
''        Do While Not rsTemp.EOF
''            SQLQ = "SELECT * FROM HREMPSKL WHERE SE_EMPNBR = " & xEmpNo & " "
''            SQLQ = SQLQ & "AND SE_SKILL = '" & rsTemp("JS_SKILL") & "' " '
''            SQLQ = SQLQ & "AND SE_LEVEL = " & rsTemp("JS_EXPFACT") & " "
''            If rsEmpSki.State <> 0 Then rsEmpSki.Close
''            rsEmpSki.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''            If rsEmpSki.EOF Then
''                rsEmpSki.AddNew
''                rsEmpSki("SE_COMPNO") = "001"
''                rsEmpSki("SE_EMPNBR") = xEmpNo
''                rsEmpSki("SE_SKILL") = rsTemp("JS_SKILL")
''                rsEmpSki("SE_LEVEL") = rsTemp("JS_EXPFACT")
''                rsEmpSki("SE_DATE") = xStartDate
''                rsEmpSki("SE_LDATE") = Date
''                rsEmpSki("SE_LTIME") = Time$
''                rsEmpSki("SE_LUSER") = glbUserID
''            End If
''            rsEmpSki.Update
''            rsTemp.MoveNext
''        Loop
''        rsMain.MoveNext
''    Loop
''    rsMain.Close
''End Sub

Private Sub WFC_PT_PenCheck(Optional NewHire = "N") 'Ticket #22991 Franks 12/24/2012
Dim rsTmpEmp As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim rsTmpPos As New ADODB.Recordset
Dim SQLQ As String
Dim xMsg1 As String
Dim xMsg2 As String
Dim xMsg3 As String
Dim xMsg4 As String
Dim xMsg5 As String
'Condition:  Woodbridge needs to send part time employee info to NGS.
'If a PT employee works 20 hours or more per week, they qualify for life insurance.
'If a PT employee works 30+ hours per week, they are treated like a ft employee and will enroll into benefits.
    
    xMsg1 = "This employee qualifies for NGS Life Benefits. " & Chr(10) & "Please enter the NGS Start Date"
    xMsg3 = "This employee qualifies for all NGS Benefits." & Chr(10) & "The employee will need to log into NGS to enroll into the benefits"
    xMsg3 = xMsg3 & Chr(10) & "Please enter the NGS Start Date"
    'Ticket #23117 Franks 01/28/2013
    xMsg5 = "If this employee is working an average of 20 hours or more per week, they qualify for life insurance or if the employee average hours per week exceed 30, "
    xMsg5 = xMsg5 & "NGS will need to be manually notified of the employee's qualification for health and dental and/or life."
    xMsg5 = xMsg5 & "Info:HR does not transfer this employee information to NGS for PT employees."
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    If rsTmpEmp.State <> 0 Then rsTmpEmp.Close
    rsTmpEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmpEmp.EOF Then
        'Ticket #30359 Franks 07/11/2017 - begin
        IsWFC_CONP = False
        If Not IsNull(rsTmpEmp("ED_EMP")) Then
            If rsTmpEmp("ED_EMP") = "CONP" Then
                IsWFC_CONP = True
            End If
        End If
        If IsWFC_CONP Then
            clpJob.Enabled = False
        End If
        'Ticket #30359 Franks 07/11/2017 - end
        
        If rsTmpEmp("ED_WORKCOUNTRY") = "U.S.A." And rsTmpEmp("ED_PT") = "PT" And Not rsTmpEmp("ED_DIV") = "1094" Then
            'check NGS Eligible and Start Date
            SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1,ER_OTHERDATE2 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
            rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpOther.EOF Then
                SQLQ = "SELECT * FROM HR_JOB_HISTORY Where JH_EMPNBR = " & glbLEE_ID & " "
                SQLQ = SQLQ & "AND NOT JH_CURRENT = 0 "
                If rsTmpPos.State <> 0 Then rsTmpPos.Close
                rsTmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTmpPos.EOF Then
                    If NewHire = "N" Then
                        'if the Hours per week is between 20 and 29.99
                        If rsTmpPos("JH_WHRS") >= 20 And rsTmpPos("JH_WHRS") < 30 Then
                            If IsNull(rsEmpOther("ER_OTHERDATE1")) Then 'NGS Start Date
                                    MsgBox xMsg1: Exit Sub
                            End If
                        End If
                        'if the Hours per week is week is 30 or more hours
                        If rsTmpPos("JH_WHRS") >= 30 Then
                            If IsNull(rsEmpOther("ER_OTHERDATE1")) Then 'NGS Start Date
                                    MsgBox xMsg3: Exit Sub
                            End If
                        End If
                    Else
                        'If rsTmpPos("JH_WHRS") >= 20 Then
                            'Ticket #23575 Franks 04/12/2013 - Remove from program
                            'MsgBox xMsg5: Exit Sub
                        'End If
                    End If
                End If
                rsTmpPos.Close
            End If
            rsEmpOther.Close
        End If
    End If
    rsTmpEmp.Close
End Sub

Private Sub WFC_PT_PenChanged() 'Ticket #22991 Franks 12/24/2012
Dim rsTmpEmp As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim rsTmpPos As New ADODB.Recordset
Dim xOldWhrs, xNewWhrs
Dim SQLQ As String
Dim xMsg1 As String
Dim xMsg2 As String
Dim xMsg3 As String
Dim xMsg4 As String
    If fglbNew Then Exit Sub 'for change only
    If Not IsNumeric(oWHRS) Then Exit Sub
    If Not IsNumeric(medHours(1).Text) Then Exit Sub
    If Round(oWHRS, 2) = Round(medHours(1).Text, 2) Then
        Exit Sub 'no change
    End If
    xOldWhrs = oWHRS
    xNewWhrs = medHours(1).Text
    
    xMsg1 = "This employee qualifies for NGS Life Benefits. " & Chr(10) & "Please enter the new NGS Start Date"
    xMsg2 = "This employee no longer qualifies for NGS Benefits." & Chr(10) & "Please enter the NGS End Date"
    xMsg3 = "This employee qualifies for all NGS Benefits." & Chr(10) & "The employee will need to log into NGS to enroll into the benefits."
    xMsg3 = xMsg3 & Chr(10) & "Please enter the new NGS Start Date"
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    If rsTmpEmp.State <> 0 Then rsTmpEmp.Close
    rsTmpEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmpEmp.EOF Then
        If rsTmpEmp("ED_WORKCOUNTRY") = "U.S.A." And rsTmpEmp("ED_PT") = "PT" And Not rsTmpEmp("ED_DIV") = "1094" Then
            'check NGS Eligible and Start Date
            SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1,ER_OTHERDATE2 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
            rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpOther.EOF Then
                If chkCurrent(0).Value Then 'currect position
                    If xOldWhrs >= 20 And xOldWhrs < 30 And xNewWhrs >= 30 Then
                        'o   If the weekly hours increase from 20-20.99 to 30 or greater, pop-up a message saying
                        MsgBox xMsg3
                        Exit Sub
                    End If
                    If xOldWhrs >= 30 And xNewWhrs >= 20 And xNewWhrs < 30 Then
                        'o   If the weekly hours decrease from 30 or greater to 20-20.99 to, pop-up a message saying
                        MsgBox xMsg1
                        Exit Sub
                    End If
                    If xOldWhrs >= 20 And xNewWhrs < 20 Then
                        'o   If the weekly hours decrease from 30 or greater or 20-20.99 to under 20 hours, pop-up a message saying
                        MsgBox xMsg2
                        Exit Sub
                    End If
                End If
                
                SQLQ = "SELECT * FROM HR_JOB_HISTORY Where JH_EMPNBR = " & glbLEE_ID & " "
                SQLQ = SQLQ & "AND NOT JH_CURRENT = 0 "
                If rsTmpPos.State <> 0 Then rsTmpPos.Close
                rsTmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTmpPos.EOF Then
                    'if the Hours per week is between 20 and 29.99
                    If rsTmpPos("JH_WHRS") >= 20 And rsTmpPos("JH_WHRS") < 30 Then
                        If IsNull(rsEmpOther("ER_OTHERDATE1")) Then 'NGS Start Date
                                MsgBox xMsg1: Exit Sub
                        Else
                            If rsTmpEmp("ED_EMPTYPE") = "N" Then
                                MsgBox xMsg2: Exit Sub
                            End If
                        End If
                    End If
                    'if the Hours per week is week is 30 or more hours
                    If rsTmpPos("JH_WHRS") >= 30 Then
                        If IsNull(rsEmpOther("ER_OTHERDATE1")) Then 'NGS Start Date
                                MsgBox xMsg3: Exit Sub
                        Else
                            If rsTmpEmp("ED_EMPTYPE") = "N" Then
                                MsgBox xMsg4: Exit Sub
                            End If
                        End If
                    End If
                End If
                rsTmpPos.Close
            End If
            rsEmpOther.Close
        End If
    End If
    rsTmpEmp.Close
End Sub

Private Sub Update_Related_SalaryPerformance_History(oPosCode, oPosStartDt)
    Dim rsSalHis As New ADODB.Recordset
    Dim rsPerfHis As New ADODB.Recordset
    Dim SQLQ As String
    
    'Retrieve and Update Salary record matching the original Position and Start Date
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND SH_JOB = '" & oPosCode & "'"
    SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(oPosStartDt)
    rsSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSalHis.EOF
        rsSalHis("SH_JOB") = clpJob.Text
        rsSalHis("SH_SDATE") = dlpStartDate.Text
        rsSalHis("SH_LDATE") = Date
        rsSalHis("SH_LTIME") = Time$
        rsSalHis("SH_LUSER") = glbUserID
        rsSalHis.Update
        
        rsSalHis.MoveNext
    Loop
    rsSalHis.Close
    Set rsSalHis = Nothing

    'Retrieve and Update Performance record matching the original Position and Position History's Job ID
    SQLQ = "SELECT * FROM HR_PERFORM_HISTORY"
    SQLQ = SQLQ & " WHERE PH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND PH_JOB = '" & oPosCode & "'"
    SQLQ = SQLQ & " AND PH_JOB_ID = " & Data1.Recordset("JH_ID")
    rsPerfHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsPerfHis.EOF
        rsPerfHis("PH_JOB") = clpJob.Text
        rsPerfHis("PH_LDATE") = Date
        rsPerfHis("PH_LTIME") = Time$
        rsPerfHis("PH_LUSER") = glbUserID
        rsPerfHis.Update
        
        rsPerfHis.MoveNext
    Loop
    rsPerfHis.Close
    Set rsPerfHis = Nothing

End Sub

Private Sub WFCHRSoftDispValues()
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp

If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsCanid.EOF Then
        Exit Sub
    End If

    'Position info -------------- begin
    If Not IsNull(rsCanid("SF_POSITIONCODE")) Then clpJob.Text = rsCanid("SF_POSITIONCODE")
    If Not IsNull(rsCanid("SF_STARTDATE")) Then dlpStartDate.Text = rsCanid("SF_STARTDATE")
    rsCanid.Close
    
    Call WFCReptDisp 'Ticket #29343 Franks 11/01/2016
End If

End Sub

Private Function isValidWFCJob(xJob, xEmpNo) 'Ticket #24767 Franks 12/10/2013
'"   If the employee's union code is not NONE or EXEC and their Category is FT, the Position Status must say "REG".
'"   If the employee's union code is NONE or EXEC and their Category is FT, the Position Status must not say "REG".
'Ignore the logic for any positions beginning with "EX". Those are executive positions that don't follow our rule. - Jerry
Dim rsLocEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xStatus
Dim retVal 'As Boolean
    retVal = 0
    If Not glbtermopen Then 'check active employees only
        If Left(xJob, 2) = "EX" Then
            'don't check
        Else
            SQLQ = "SELECT ED_EMPNBR, ED_ORG, ED_PT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsLocEmp.EOF Then
                If Not IsNull(rsLocEmp("ED_PT")) And Not IsNull(rsLocEmp("ED_ORG")) Then
                    'find the Job Status
                    xStatus = getPosMasterValueByField(xJob, "JB_STATUS")
                    If Not (rsLocEmp("ED_ORG") = "NONE" Or rsLocEmp("ED_ORG") = "EXEC") Then
                        If rsLocEmp("ED_PT") = "FT" Then
                            If Not xStatus = "REG" Then
                                retVal = 1
                            End If
                        End If
                    End If
                    If (rsLocEmp("ED_ORG") = "NONE" Or rsLocEmp("ED_ORG") = "EXEC") Then
                        If rsLocEmp("ED_PT") = "FT" Then
                            If xStatus = "REG" Then
                                retVal = 2
                            End If
                        End If
                    End If
                End If
            End If
            rsLocEmp.Close
        End If
    End If
    isValidWFCJob = retVal
End Function

Private Sub Populate_ComShift()
    'Vitalaire
    'Ticket #24976 - Label changed, and add dropdown list
    If glbCompSerial = "S/N - 2380W" Then
        comShift.Width = 3500
        comShift.Clear
        comShift.AddItem "A - Operations- Cylinder & Bulk Centers"
        comShift.AddItem "B - Operations- Plants & Pipelines"
        comShift.AddItem "C - Installations"
        comShift.AddItem "D - Maintenance & Services"
        comShift.AddItem "E - Logistics"
        comShift.AddItem "F - Manufacturing"
        comShift.AddItem "G - R & D"
        comShift.AddItem "H - Technical/ Industrial Support"
        comShift.AddItem "I - Project Management"
        comShift.AddItem "J - Sales & Business Development"
        comShift.AddItem "K - Sales Administration/ Active Cycle"
        comShift.AddItem "L - Marketing"
        comShift.AddItem "M - Finance & Administration"
        comShift.AddItem "N - Human Resources & Employee Relations"
        comShift.AddItem "O - Others"
        comShift.AddItem "P - IT & Telecom"
    End If
    'Ticket #25952 Franks 11/03/2014
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc
        comShift.Width = 1800
        comShift.Clear
        comShift.AddItem "D - Day Shift"
        comShift.AddItem "A - Afternoon Shift"
        comShift.AddItem "R - Rotating"
        comShift.AddItem "T - Afternoon2" 'Ticket #29272 Franks 09/26/2016
    End If
End Sub

Private Function GetComShiftIndex(xSHIFT)
    Dim xIndex As Integer
    xIndex = -1

    'Vitalaire
    'Ticket #24976 - Label changed, and add dropdown list
    If glbCompSerial = "S/N - 2380W" Then
        Select Case xSHIFT
            Case "A": xIndex = 0
            Case "B": xIndex = 1
            Case "C": xIndex = 2
            Case "D": xIndex = 3
            Case "E": xIndex = 4
            Case "F": xIndex = 5
            Case "G": xIndex = 6
            Case "H": xIndex = 7
            Case "I": xIndex = 8
            Case "J": xIndex = 9
            Case "K": xIndex = 10
            Case "L": xIndex = 11
            Case "M": xIndex = 12
            Case "N": xIndex = 13
            Case "O": xIndex = 14
            Case "P": xIndex = 15
        End Select
    End If
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
        Select Case xSHIFT
            Case "D": xIndex = 0
            Case "A": xIndex = 1
            Case "R": xIndex = 2
            Case "T": xIndex = 3
        End Select
    End If
    
    GetComShiftIndex = xIndex
End Function

Private Sub WFCDefaultHours()
Dim xHrsDay, xHrsWk, xHrsPay
    'Ticket #25221 Franks 03/17/2014 - BEGIN
    If NewHireForms.count > 0 Then 'new hire
        If (glbUNION = "NONE" Or glbUNION = "EXEC" Or glbUNION = "-NON" Or glbUNION = "-EXE") Then  'salaried
            xHrsDay = 8
            xHrsWk = 40
            xHrsPay = 86.67
            If IsNumeric(glbTrsHourWeek) Then
                If glbTrsHourWeek > 0 Then
                    xHrsWk = glbTrsHourWeek
                End If
            End If
          Else  'hourly
            xHrsDay = 8
            xHrsWk = 40
            xHrsPay = 40
            If IsNumeric(glbTrsHourWeek) Then
                If glbTrsHourWeek > 0 Then
                    xHrsWk = glbTrsHourWeek
                    xHrsPay = glbTrsHourWeek
                End If
            End If
        End If
    Else 'non new hire: get previous hours
        xHrsDay = GetPrePositionData(glbLEE_ID, "JH_DHRS", "")
        xHrsWk = GetPrePositionData(glbLEE_ID, "JH_WHRS", "")
        xHrsPay = GetPrePositionData(glbLEE_ID, "JH_PHRS", "")
    End If
    medHours(0).Text = xHrsDay
    medHours(1).Text = xHrsWk
    medHours(2).Text = xHrsPay
    'Ticket #25221 Franks 03/17/2014 - END
    
    ''''Ticket #24337 Franks 09/23/2013 - begin
    '''If (glbUNION = "NONE" Or glbUNION = "EXEC") Then  'salaried
    '''Ticket #24184 Franks 10/17/2013
    ''If (glbUNION = "NONE" Or glbUNION = "EXEC" Or glbUNION = "-NON" Or glbUNION = "-EXE") Then  'salaried
    ''    medHours(0).Text = 8
    ''    medHours(1).Text = 40
    ''    medHours(2).Text = 86.67
    ''Else 'hourly
    ''    medHours(0).Text = 8
    ''    medHours(1).Text = 40
    ''    medHours(2).Text = 40
    ''End If
    '''Ticket #24337 Franks 09/23/2013 - end.
    
    'Ticket #25884 Franks 08/19/2014 - begin
    'glbPlantCode
    txtReptAuthority(3).Text = getWFCRA4(glbEmpDiv)
    'Ticket #25884 Franks 08/19/2014 - end
End Sub

Private Function IsFirstEmpPosition(xEmpNo)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = False
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rs.EOF Then
        retVal = True
    End If
    rs.Close
    IsFirstEmpPosition = retVal
End Function

Private Function getWaltersIncEmailBody() 'Walters Inc 'Ticket #25952 Franks 11/03/2014
Dim retVal
    retVal = ""
    If fglbNew Then 'new record
        If NewHireForms.count > 0 Then ' new hire
            retVal = retVal & "This will serve to confirm that the following employee's position title has been created" & vbCrLf & vbCrLf
        Else
            retVal = retVal & "This will serve to confirm that the following employee's position title has been changed" & vbCrLf & vbCrLf
        End If
            retVal = retVal & "Name: " & lblEEName.Caption & vbCrLf
            If Len(savJOB) > 0 And Not savJOB = clpJob.Text Then
                retVal = retVal & "Previous position: " & getPosDesc(savJOB) & vbCrLf
            End If
            retVal = retVal & "New position: " & getPosDesc(clpJob.Text) & vbCrLf
            retVal = retVal & "Effective Date: " & dlpStartDate.Text & vbCrLf
            retVal = retVal & "Shift: " & Mid(comShift.Text, 5, 15) & vbCrLf
    End If
    getWaltersIncEmailBody = retVal
End Function

Private Sub VitalAireJobFamilyScreen() 'Ticket #26233 Franks 11/24/2014 VitalAire Canada Inc.
    lblBand.Top = 4440
    lblBand.Left = 6150
    clpCode(6).Top = 4440
    clpCode(6).Left = 7410
    frmVitalAireJobFamily.Left = 6100
    frmVitalAireJobFamily.Top = 4770
    frmVitalAireJobFamily.BorderStyle = 0
    frmVitalAireJobFamily.Visible = True
End Sub

Private Sub VitalAireJobFamilyDesc(JobCode) 'Ticket #26233 Franks 11/24/2014 VitalAire Canada Inc.
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JB_CODE,JB_JOBFAMILY,JB_SUBJOBFAMILY,JB_JOBFAMILYGRP FROM HRJOB WHERE JB_CODE='" & JobCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB("JB_JOBFAMILY")) Then
            txtDouDiv(0).Text = rsJOB("JB_JOBFAMILY")
        Else
            txtDouDiv(0).Text = ""
            lblDouDivDesc(0).Caption = ""
        End If
        If Not IsNull(rsJOB("JB_SUBJOBFAMILY")) Then
            txtDouDiv(1).Text = rsJOB("JB_SUBJOBFAMILY")
        Else
            txtDouDiv(1).Text = ""
            lblDouDivDesc(1).Caption = ""
        End If
        If Not IsNull(rsJOB("JB_JOBFAMILYGRP")) Then
            txtDouDiv(2).Text = rsJOB("JB_JOBFAMILYGRP")
        Else
            txtDouDiv(2).Text = ""
            lblDouDivDesc(2).Caption = ""
        End If
    End If
    rsJOB.Close
End Sub

'Private Sub NYCHScreenSetup() 'Ticket #26979 Franks 04/24/2015
'    frmNYCH.Top = frmLinamar(0).Top - 30
'    frmNYCH.Left = 0
'    frmNYCH.BorderStyle = 0
'    frmNYCH.Width = 5500
'    frmNYCH.Visible = True
'
'    lblSalDist.Caption = lStr("Salary Distribution") ' "Distribution Code" ' "Accpac Distribution Code"
'    lblTitle(29).Left = 60
'    clpSalDist.Left = 1825
'    txtLabel(3).DataField = ""
'    txtUSRLABEL3.DataField = "JH_USRLABEL3"
'End Sub

'Private Sub clpSalDist_LostFocus() 'Ticket #26979 Franks 04/24/2015
'    txtUSRLABEL3.Text = clpSalDist.Text
'End Sub

'Private Sub txtUSRLABEL3_Change() 'Ticket #26979 Franks 04/24/2015
'    clpSalDist.Text = txtUSRLABEL3.Text
'End Sub

Private Function AUDITSALY(empNo, NSalary, oPayP, oJob, oGrid, oPayrollID, OSalCD, oHrsWk, OEDate, ONDate, OReason)
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim TB As New ADODB.Recordset
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITSALY = False


TB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & empNo, gdbAdoIhr001, adOpenForwardOnly
If Not TB.EOF Then
    If IsNull(TB("ED_PT")) Then
        xPT = ""
    Else
        xPT = TB("ED_PT")
    End If
    If IsNull(TB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = TB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
TB.Close
'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'strfields added by Bryan 02/Dec/05 TICKET#9899
strFields = "AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL,AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_PAYP, "
strFields = strFields & "AU_OLDPAYP, AU_OLDPAYP, AU_JOB, AU_GRID, AU_PAYROLL_ID, AU_SALCD, AU_WHRS, AU_SEDATE, AU_SNDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_SREASON "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

'If OSalary <> NSalary Then GoTo MODUPD
''If OPayp <> NPayp Then GoTo MODUPD      'laura jan 28, 1998
'If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD

'GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDIV
TA("AU_SALARY") = NSalary
'TA("AU_OLDSAL") = NSalary  'Ticket #27056 - Do not save this as it will not export to Payroll if Salary = OldSal
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_JOB") = oJob          ' FRANK 4/5/2000
TA("AU_GRID") = oGrid
If glbMulti Then TA("AU_PAYROLL_ID") = oPayrollID
TA("AU_SALCD") = OSalCD
TA("AU_WHRS") = oHrsWk 'ADDED BY RAUBREY 7/7/97
'If OEDate <> NEDate Then TA("AU_SEDATE") = IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
'If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99
TA("AU_SEDATE") = OEDate   'IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
TA("AU_SNDATE") = ONDate   'IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_SREASON") = OReason

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = empNo

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    'TA("AU_LDATE") = Format(DateAdd("d", 14, NEDate), "SHORT DATE")
    TA("AU_LDATE") = Format(DateAdd("d", 14, OEDate), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        'If CVDate(NEDate) > CVDate(Date) Then
        If CVDate(OEDate) > CVDate(Date) Then
            'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
            TA("AU_LDATE") = Format(OEDate, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
        TA("AU_LDATE") = Format(OEDate, "SHORT DATE")
    End If
End If
'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")

TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & empNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then TA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If
TA.Update


MODNOUPD:
AUDITSALY = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "SAME SALARY UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
End Function

Private Sub DispJobCode() 'Ticket #27774 Franks 11/18/2015
    lblJob.Top = lblPosTitle.Top - 360
    lblJob.Left = lblPosTitle.Left
    lblJob.Visible = True
    lblJobDesc.Top = lblPosTitle.Top - 360
    lblJobDesc.Left = txtShift.Left
    lblJobDesc.Visible = True
End Sub

Private Sub ReptsEffDatesScreen()
    fraReptEDate.Top = 1080
    fraReptEDate.Left = frmJobEnd.Left  '5000 '5880
    fraReptEDate.Width = 4335
    fraReptEDate.Height = 1400
    fraReptEDate.BorderStyle = 0
    fraReptEDate.Visible = True
    dlpRptDate(1).DataField = "JH_EDATEREPT1"
    dlpRptDate(2).DataField = "JH_EDATEREPT2"
    dlpRptDate(3).DataField = "JH_EDATEREPT3"
    dlpRptDate(4).DataField = "JH_EDATEREPT4"
    fraPosition.Left = frmJobEnd.Left
    'optSalary(2).Visible = True
    If (glbUNION = "NONE" Or glbUNION = "EXEC") Then
        optSalary(2).Visible = True
        fraPosition.Width = 5500
    Else
        optSalary(2).Visible = False
        fraPosition.Width = 3225
    End If
    
    If glbWFC Then 'Ticket #29438 Franks 11/07/2016
        imgPosFilled(0).Top = clpJob.Top
        imgPosFilled(0).Visible = True
        imgPosFilled(1).Visible = True
        If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
            Command1.Visible = True
        End If
    End If
End Sub

Private Sub WFC_fraPosition() 'Ticket #29343 Franks 10/18/2016
'    If (glbUNION = "NONE" Or glbUNION = "EXEC") Then
'        'optSalary(2).Visible = True
'        fraPosition.Width = 5500
'    Else
'        'optSalary(2).Visible = False
'        'fraPosition.Width = 3225
'    End If
End Sub

Private Sub NiagaraFallsScreen() 'City of Niagara Falls Ticket #27681 Franks 12/10/2015
    fraPosition.Top = 345 + 450
    lblEStatus.Top = 345
    lblEStatus.Left = 7380
    lblEStatus.Visible = True
    clpCode(7).Top = 345
    clpCode(7).Left = 9240
    clpCode(7).Visible = True
    clpCode(7).DataField = "JH_ESTATUS"
End Sub

Private Function PrimaryPositionExists(xEmpnbr, Optional xJH_ID)
Dim rsEmpJob As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean

    'Check if Primary Position already exists
    
    retVal = False
    
    SQLQ = "SELECT JH_EMPNBR, JH_ID, JH_PRIMARY FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpnbr & " "
    SQLQ = SQLQ & " AND JH_PRIMARY <> 0 "
    If Not IsMissing(xJH_ID) Then
        SQLQ = SQLQ & " AND NOT JH_ID = " & xJH_ID & " "
    End If
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpJob.EOF Then
        retVal = True
    End If
    PrimaryPositionExists = retVal

End Function

Private Function WFCNewHireEmailSending()
Dim rsEmp As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUserEmpNo
Dim xEmpPlant
Dim xStr
Dim xFName

    xUserEmpNo = GetUserEmpNo(glbUserID)
    If Len(xUserEmpNo) > 0 Then
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xUserEmpNo & " "
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            xFName = GetEmpData(glbLEE_ID, "ED_FNAME")
            MailBodyN = rsEmp("ED_FNAME") & " " & rsEmp("ED_SURNAME") & " from " & GetTABLCodePub("EDSE", rsEmp("ED_SECTION")) & " hired "
            MailBodyN = MailBodyN & "Employee #" & glbLEE_ID '& vbCrLf
            MailBodyN = MailBodyN & " - " & xFName & " " & GetEmpData(glbLEE_ID, "ED_SURNAME") & " as "
            MailBodyN = MailBodyN & clpJob.Caption & vbCrLf & vbCrLf  '", "
            'Ticket #28895 Franks 07/25/2016 - begin 'add department and Reporting Authority #1's
            xStr = GetEmpData(glbLEE_ID, "ED_DEPTNO")
            If Len(xStr) > 0 Then
                MailBodyN = MailBodyN & xFName & " will work in Department " & GetDeptName(xStr, "DF_NAME") & " " '
                If Len(elpReptAuthShow(0).Text) > 0 Then
                    MailBodyN = MailBodyN & "and will be Reporting to " & elpReptAuthShow(0).Caption & " "
                End If
                MailBodyN = MailBodyN & vbCrLf & vbCrLf
            End If
            'Ticket #28895 Franks 07/25/2016 - end
            MailBodyN = MailBodyN & "The Effective Date of Hire is " & GetEmpData(glbLEE_ID, "ED_DOH") & vbCrLf & vbCrLf
            
            ''xEmpPlant = GetEmpData(glbLEE_ID, "ED_SECTION")
            ''If Not IsNull(xEmpPlant) Then
            ''    If xEmpPlant = "MISS" Or xEmpPlant = "TROY" Then
            ''        'get Network Login ID
            ''        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & " "
            ''        If rsEmpOther.State <> 0 Then rsEmpOther.Close
            ''        rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            ''        If Not rsEmpOther.EOF Then
            ''            If Not IsNull(rsEmpOther("ER_NETWORKLOGIN")) Then
            ''                If Len(Trim(rsEmpOther("ER_NETWORKLOGIN"))) > 0 Then
            ''                    MailBodyN = MailBodyN & "Network Login is " & Trim(rsEmpOther("ER_NETWORKLOGIN")) & vbCrLf & vbCrLf
            ''                End If
            ''            End If
            ''        End If
            ''        rsEmpOther.Close
            ''    End If
            ''End If
            'Ticket #30472 Franks 08/11/2017 - begin
            xStr = get_EmpOtherByField(glbLEE_ID, "ER_NETWORKLOGIN")
            If Not IsNull(xStr) Then
                If Left(xStr, 13) = "Network Login" Then
                Else
                    xStr = Trim(xStr)
                    If Len(xStr) > 0 Then
                        MailBodyN = MailBodyN & "Network Login is " & xStr & vbCrLf & vbCrLf
                    End If
                End If
            End If
            'Ticket #30472 Franks 08/11/2017 - end
            
            xStr = GetEmpData(glbLEE_ID, "ED_EMAIL")
            If Len(xStr) > 0 Then
                MailBodyN = MailBodyN & xFName & "'s email address is " & xStr & vbCrLf & vbCrLf
            End If
            
            Call imgEmail_NewHire
        End If
    End If
End Function

Private Sub LinamarSceenSetup()
Dim X
    For X = 0 To 4
        frmLinamar(X).Visible = True
    Next
    lblHrsDay.FontBold = True
    lblHrsWeek.FontBold = True
    lblHrsPayPeriod.FontBold = True
    lblShift.FontBold = True
    cmdEditLable.Visible = True
    
    'Ticket #28846 Franks 07/14/2016 - begin
    txtShift.Visible = False
    txtShift.DataField = ""
    lblReptAuth(0).FontBold = True
    clpCode(8).Top = txtShift.Top
    clpCode(8).Left = clpJob.Left
    clpCode(8).Visible = True
    'Ticket #28846 Franks 07/14/2016 - end
    
    'Ticket #28846 Franks 08/16/2016
    clpCode(3).MaxLength = 10
    clpCode(3).TextBoxWidth = 1000
    
    frmLinamar(1).Top = 5730
    frmLinamar(1).Left = 0
    frmLinamar(2).Top = 6240
    frmLinamar(2).Left = 0
    frmLinamar(3).Top = 6780
    frmLinamar(3).Left = 0
    frmLinamar(4).Top = 7410
    frmLinamar(4).Left = 0
    
    frmWFCDIV.Top = 130
    frmWFCDIV.BorderStyle = 0
End Sub

Sub getCodes() 'Ticket #28846 Franks 07/14/2016
If rsDATA.EOF Then Exit Sub
If glbLinamar Then
    If Not IsNull(rsDATA("JH_SHIFT")) Then
        clpCode(8).Text = Mid(rsDATA("JH_SHIFT"), 4)
    Else
        clpCode(8).Text = ""
    End If
End If
End Sub

Private Sub UpdCodes() 'Ticket #28846 Franks 07/14/2016
    If glbLinamar Then
        If Trim(clpCode(8).Text) <> "" Then
            rsDATA("JH_SHIFT") = getShiftCodeforLinamar(clpCode(8).TransDiv & clpCode(8).Text)
        Else
            rsDATA("JH_SHIFT") = ""
        End If
    End If
End Sub
Private Function getReptPosEmpListByEmp(xEmpNo) 'Ticket #29438 Franks 11/07/2016
Dim rsEmp As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim rsJobHis As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDIV, xDeptno, xGLNO, xPosCtrl
Dim xSec, xCunt
Dim xBudgNo, xVacantNo, I
Dim xReptPosCode, xReptPosDesc, xPosCodeDesc, xPosCode
Dim retVal
    retVal = ""
    If Len(xEmpNo) = 0 Then
        getReptPosEmpListByEmp = retVal
        Exit Function
    End If
    If Not IsNumeric(xEmpNo) Then ' Len(xPosCode) = 0 Then
        getReptPosEmpListByEmp = retVal
        Exit Function
    End If
    
    'Exit Function '
    
    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
    If Len(xEmpNo) > 0 Then
        If IsNumeric(xEmpNo) Then
            SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        End If
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        getReptPosEmpListByEmp = retVal
        Exit Function
    End If
    If Not rsEmp.EOF Then
        xPosCode = rsEmp("JH_JOB")
    End If
    
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xPosCode & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'xReptPosCode = xPosCode '""
    xPosCodeDesc = ""
    If rsJOB.EOF Then
        getReptPosEmpListByEmp = retVal
        Exit Function
    Else
        xPosCodeDesc = rsJOB("JB_DESCR")
    End If
    
    xCunt = 0
    
    '????? - NOT DONE YET
    
    SQLQ = "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 " 'AND JH_JOB = '" & xReptPosCode & "' "
    If Len(xEmpNo) > 0 Then
        If IsNumeric(xEmpNo) Then
            'SQLQ = SQLQ & "AND JH_REPTAU = " & xEmpNo & " "
            SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        End If
    End If
    rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsJobHis.EOF
        rsEListWRK.AddNew
        rsEListWRK("TT_COMPNO") = "001"
        rsEListWRK("TT_EMPNBR") = rsJobHis("JH_EMPNBR")
        rsEListWRK("TT_SURNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_SURNAME")
        rsEListWRK("TT_FNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_FNAME")
        rsEListWRK("TT_WRKEMP") = glbUserID
        rsEListWRK.Update
        xCunt = xCunt + 1

        rsJobHis.MoveNext
    Loop
    rsJobHis.Close

    getReptPosEmpListByEmp = retVal
 
    Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "RZEmpList4.rpt" '"RZEmpList3.rpt"
    Me.vbxCrystal2.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal2.Formulas(0) = "rTitle='" & lblReptAuth(0).Caption & " information'"
    Me.vbxCrystal2.Connect = RptODBC_SQL
    'window title if appropriate
    Me.vbxCrystal2.WindowTitle = lblReptAuth(0).Caption & " Employee Position Information"
    Me.vbxCrystal2.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal2.Action = 1
    vbxCrystal2.Reset


End Function

Private Function getReptPosEmpListByPos(xPosCode, xEmpNo) 'Ticket #29438 Franks 11/07/2016
Dim snapJobCount As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim rsJobHis As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDIV, xDeptno, xGLNO, xPosCtrl
Dim xSec, xCunt
Dim xBudgNo, xVacantNo, I
Dim xReptPosCode, xReptPosDesc, xPosCodeDesc
Dim retVal
    retVal = ""
    If Len(xPosCode) = 0 Then
        getReptPosEmpListByPos = retVal
        Exit Function
    End If
    
    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    

    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xPosCode & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'xReptPosCode = xPosCode '""
    xPosCodeDesc = ""
    If rsJOB.EOF Then
        getReptPosEmpListByPos = retVal
        Exit Function
    End If
    
    xCunt = 0
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB("JB_REPTAU")) Then
            If Len(rsJOB("JB_REPTAU")) > 0 Then
                xReptPosCode = rsJOB("JB_REPTAU")
                xPosCodeDesc = rsJOB("JB_DESCR")
                
                SQLQ = "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB = '" & xReptPosCode & "' "
                If Len(xEmpNo) > 0 Then
                    If IsNumeric(xEmpNo) Then
                        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
                    End If
                End If
                rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If rsJobHis.EOF Then
                    ''rsEListWRK.AddNew
                    ''rsEListWRK("TT_COMPNO") = "001"
                    ''rsEListWRK("TT_EMPNBR") = 0 ' rsJobHis("JH_EMPNBR")
                    ''rsEListWRK("TT_SURNAME") = "Vacant " 'GetEmpData(rsJobHis("JH_EMPNBR"), "ED_SURNAME")
                    ''rsEListWRK("TT_FNAME") = "" ' GetEmpData(rsJobHis("JH_EMPNBR"), "ED_FNAME")
                    ''rsEListWRK("TT_WRKEMP") = glbUserID
                    ''rsEListWRK.Update
                    ''xCunt = xCunt + 1
                Else
                    Do While Not rsJobHis.EOF
                        rsEListWRK.AddNew
                        rsEListWRK("TT_COMPNO") = "001"
                        rsEListWRK("TT_EMPNBR") = rsJobHis("JH_EMPNBR")
                        rsEListWRK("TT_SURNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_SURNAME")
                        rsEListWRK("TT_FNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_FNAME")
                        rsEListWRK("TT_WRKEMP") = glbUserID
                        rsEListWRK.Update
                        xCunt = xCunt + 1
    
                        rsJobHis.MoveNext
                    Loop
                End If
                rsJobHis.Close
                
            End If
        End If
    End If


    getReptPosEmpListByPos = retVal
 
    Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "RZEmpList4.rpt" '"RZEmpList3.rpt"
    Me.vbxCrystal2.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal2.Formulas(0) = "rTitle='Rept Position of " & xPosCode & "(" & xPosCodeDesc & ")" & "'"
    Me.vbxCrystal2.Connect = RptODBC_SQL
    'window title if appropriate
    Me.vbxCrystal2.WindowTitle = "Rept. Authority Position Employees List Report"
    Me.vbxCrystal2.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal2.Action = 1
    vbxCrystal2.Reset

    
End Function

Private Sub WFC_CONP_Fields()
Dim xDIV
    xDIV = GetEmpData(glbLEE_ID, "ED_DIV")
    clpJob.Text = getWFC_CONP_Pos(xDIV)
    medHours(0) = 0
    medHours(1) = 0
    medHours(2) = 0
    clpCode(1).Text = "CE"
End Sub

