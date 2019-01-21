VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUBENEFITS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Benefits"
   ClientHeight    =   8490
   ClientLeft      =   315
   ClientTop       =   885
   ClientWidth     =   8880
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optBenefit 
      Alignment       =   1  'Right Justify
      Caption         =   "Change Cost Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   14
      Top             =   3030
      Width           =   2175
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8040
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton optBenefit 
      Alignment       =   1  'Right Justify
      Caption         =   "Benefit Group "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   3030
      Width           =   1695
   End
   Begin VB.OptionButton optBenefit 
      Alignment       =   1  'Right Justify
      Caption         =   "Benefit Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   3030
      Value           =   -1  'True
      Width           =   1815
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   6120
      TabIndex        =   7
      Top             =   630
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Top             =   300
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin VB.Frame Frame1 
      Height          =   405
      Left            =   4920
      TabIndex        =   53
      Top             =   1590
      Width           =   4455
      Begin VB.CheckBox chkEmployee 
         Alignment       =   1  'Right Justify
         Caption         =   "Terminated Employees"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Width           =   1965
      End
      Begin VB.CheckBox chkEmployee 
         Alignment       =   1  'Right Justify
         Caption         =   "Active Employees"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Value           =   1  'Checked
         Width           =   1965
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   1350
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   20
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Tag             =   "00-Enter Union Code"
      Top             =   1020
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2010
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7315
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin VB.Frame frmMaster 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3975
      Left            =   60
      TabIndex        =   56
      Top             =   3300
      Width           =   9615
      Begin VB.ComboBox cmbPerOrDoll 
         Height          =   315
         ItemData        =   "fubenes.frx":0000
         Left            =   7590
         List            =   "fubenes.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "40-Select Dallor or Percentage"
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox txtCovType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1635
         MaxLength       =   1
         TabIndex        =   22
         Tag             =   "00-Type of Coverage (Single or Family)"
         Top             =   840
         Width           =   330
      End
      Begin VB.ComboBox comRndFactor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "fubenes.frx":0022
         Left            =   6330
         List            =   "fubenes.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "Rounding Factor"
         Top             =   1530
         Width           =   1215
      End
      Begin VB.ComboBox comSalDepn 
         Height          =   315
         ItemData        =   "fubenes.frx":0095
         Left            =   6330
         List            =   "fubenes.frx":009F
         TabIndex        =   23
         Text            =   "No"
         Top             =   810
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   7740
         TabIndex        =   58
         Top             =   1080
         Width           =   1155
         Begin VB.OptionButton optRound 
            Caption         =   "Nearest"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   210
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optRound 
            Caption         =   "Next"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.TextBox txtPreAftTax 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8700
         TabIndex        =   57
         Top             =   2700
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox comPreAftTax 
         Height          =   315
         ItemData        =   "fubenes.frx":00AC
         Left            =   7440
         List            =   "fubenes.frx":00B6
         TabIndex        =   37
         Tag             =   "Pre Tax/After Tax"
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox txtPer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7785
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "10-Enter number of units"
         Top             =   2400
         Width           =   870
      End
      Begin VB.TextBox txtTAXBEN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   40
         Tag             =   "00-Taxable Benefit    Y=Yes     N=No"
         Top             =   3060
         Width           =   855
      End
      Begin INFOHR_Controls.DateLookup dlpEDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Tag             =   "41-Effective Date of coverage"
         Top             =   510
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   17
         Tag             =   "01-Benefit - Code"
         Top             =   180
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNCD"
         MaxLength       =   10
      End
      Begin MSMask.MaskEdBox medPayPeriodAmount 
         Height          =   285
         Left            =   6330
         TabIndex        =   18
         Tag             =   "20-Amount charged for every pay period"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox medMaxAmnt 
         Height          =   285
         Left            =   6330
         TabIndex        =   21
         Tag             =   "20-Enter Maximum Amount"
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   "$#,##0.0000;($#,##0.0000)"
         PromptChar      =   " "
      End
      Begin Threed.SSFrame frmAP 
         Height          =   465
         Left            =   0
         TabIndex        =   59
         Top             =   1890
         Width           =   8775
         _Version        =   65536
         _ExtentX        =   15478
         _ExtentY        =   820
         _StockProps     =   14
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
         Font3D          =   1
         Begin Threed.SSOption optActual 
            Height          =   225
            Index           =   0
            Left            =   1920
            TabIndex        =   30
            Tag             =   "Choose actual or premium"
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Actual"
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
         Begin Threed.SSOption optActual 
            Height          =   225
            Index           =   1
            Left            =   4680
            TabIndex        =   31
            TabStop         =   0   'False
            Tag             =   "Choose actual or premium"
            Top             =   150
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Premium"
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
         Begin VB.Label lblAP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            DataField       =   "BF_PREMIUM"
            DataSource      =   "Data2"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   960
            TabIndex        =   60
            Top             =   210
            Visible         =   0   'False
            Width           =   435
         End
      End
      Begin MSMask.MaskEdBox medCovAmount 
         Height          =   285
         Left            =   1635
         TabIndex        =   32
         Tag             =   "20-Amount of Coverage"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$##,##0.00;($##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPPComp 
         Height          =   285
         Left            =   1635
         TabIndex        =   35
         Tag             =   "11-Percentage paid by company"
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMCCOST 
         Height          =   285
         Left            =   1635
         TabIndex        =   38
         Tag             =   "21-Monthly company cost"
         Top             =   3060
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$##,##0.00;($##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCompCost 
         Height          =   285
         Left            =   1635
         TabIndex        =   41
         Tag             =   "11-Cost of Benefit to Company"
         Top             =   3390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.0000;($#,##0.0000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medUnitCost 
         Height          =   285
         Left            =   4530
         TabIndex        =   33
         Tag             =   "20-Enter Unit Cost"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.000000;($#,##0.000000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPPE 
         Height          =   285
         Left            =   4530
         TabIndex        =   36
         Tag             =   "11-Percentage paid by employee"
         Top             =   2730
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMECOST 
         Height          =   285
         Left            =   4530
         TabIndex        =   39
         Tag             =   "21-Monthly employee cost"
         Top             =   3060
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$##,##0.00;($##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECost 
         Height          =   285
         Left            =   4530
         TabIndex        =   42
         Tag             =   "11-Cost of benefit to Employee"
         Top             =   3390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.0000;($#,##0.0000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMinCover 
         Height          =   285
         Left            =   1635
         TabIndex        =   24
         Tag             =   "20-Minimum of Coverage"
         Top             =   1215
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMaxCover 
         Height          =   285
         Left            =   6330
         TabIndex        =   25
         Tag             =   "20-Maximum of Coverage"
         Top             =   1170
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSalFactor 
         Height          =   285
         Left            =   1635
         TabIndex        =   26
         Tag             =   "20-Salary Factor"
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.0000;($#,##0.0000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTCost 
         Height          =   285
         Left            =   6945
         TabIndex        =   43
         Tag             =   "21-Total Cost of the Coverage"
         Top             =   3390
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$##,###.0000;($##,###.0000)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblBenefit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit"
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
         Left            =   0
         TabIndex        =   85
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   84
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   83
         Top             =   900
         Width           =   825
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period Amount"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4440
         TabIndex        =   82
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Amount"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   81
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage Amount"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   80
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "% Paid "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   79
         Top             =   2730
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   78
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   77
         Top             =   3390
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   76
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   660
         TabIndex        =   75
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   660
         TabIndex        =   74
         Top             =   3390
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   3240
         TabIndex        =   73
         Top             =   2430
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "% Paid Employee"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   72
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   71
         Top             =   3090
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   3240
         TabIndex        =   70
         Top             =   3420
         Width           =   825
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Per"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   6060
         TabIndex        =   69
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable Benefit"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6060
         TabIndex        =   68
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   6060
         TabIndex        =   67
         Top             =   3390
         Width           =   615
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Rounding Factor"
         Height          =   315
         Index           =   29
         Left            =   4440
         TabIndex        =   66
         Top             =   1545
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Factor"
         Height          =   315
         Index           =   27
         Left            =   0
         TabIndex        =   65
         Top             =   1605
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Coverage"
         Height          =   315
         Index           =   28
         Left            =   4440
         TabIndex        =   64
         Top             =   1185
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Coverage"
         Height          =   315
         Index           =   26
         Left            =   0
         TabIndex        =   63
         Top             =   1230
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Dependent"
         Height          =   315
         Index           =   25
         Left            =   4440
         TabIndex        =   62
         Top             =   825
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pre Tax/After Tax"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   6060
         TabIndex        =   61
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Frame frmGroup 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1065
      Left            =   -120
      TabIndex        =   86
      Top             =   3330
      Visible         =   0   'False
      Width           =   7455
      Begin INFOHR_Controls.CodeLookup clpBGroup 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   16
         Tag             =   "01-Benefit Group - Code"
         Top             =   480
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpBGroup 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Tag             =   "01-Benefit Group - Code"
         Top             =   150
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Benefit Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   89
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Benefit Group"
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
         Index           =   16
         Left            =   360
         TabIndex        =   87
         Top             =   540
         Width           =   1620
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   6120
      TabIndex        =   9
      Tag             =   "01-Benefit - Code"
      Top             =   1290
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BGMF"
      MaxLength       =   10
   End
   Begin VB.Label lblOldBGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4920
      TabIndex        =   88
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   55
      Top             =   1710
      Width           =   630
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
      Left            =   60
      TabIndex        =   54
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4920
      TabIndex        =   52
      Top             =   990
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
      Left            =   4920
      TabIndex        =   51
      Top             =   660
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
      Left            =   4920
      TabIndex        =   50
      Top             =   330
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
      TabIndex        =   49
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label textMulti 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The Union Code and FT/PT/SE/TR/OT will be validated from the Employee Basic Data"
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
      Height          =   195
      Left            =   0
      TabIndex        =   48
      Top             =   2520
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   47
      Top             =   1350
      Width           =   1350
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   46
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   45
      Top             =   690
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
      Left            =   60
      TabIndex        =   44
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmUBENEFITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SN_CASEY = "S/N - 2214W"

Dim fglbAdd%    ' it is a global add request
Dim fglbDelete%, fglbNoDept&
Dim fglbModify%, fglb_FindDept
Dim Actn
Dim ChangingFields As Boolean
Dim fglbESQLQ, fglbWSQLQ
Dim fglbSDate As Variant
Dim XUpdCount
Dim fglbDupCode As Boolean
Dim RSEMPLIST As New ADODB.Recordset 'George Mar 2,2006
Dim strEMPLIST 'George Mar 2,2006
Dim strTermEMPLIST 'George Mar 14,2006
Dim MailBody, xStr
Dim flgSendEmail  As Boolean
Dim EmpListAdd

Private Sub comPreAftTax_Change()
    If comPreAftTax = "Pre Tax" Then
        txtPreAftTax = "P"
    ElseIf comPreAftTax = "After Tax" Then
        txtPreAftTax = "A"
    Else
        txtPreAftTax = ""
    End If
End Sub

Private Sub comPreAftTax_Click()
    Call comPreAftTax_Change
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUBENEFITS"
End Sub

Private Sub medMCCOST_Change()
    If glbCompSerial = SN_CASEY And Not ChangingFields Then CalcBeneCasey
End Sub

Private Sub medMECOST_Change()
    If glbCompSerial = SN_CASEY And Not ChangingFields Then CalcBeneCasey
End Sub

Private Function AUDITBENF(ACTX)
Dim TA As New ADODB.Recordset
Dim TB As New ADODB.Recordset
Dim xPT, xDiv, xADD
Dim TC As New ADODB.Recordset
Dim SQLQ

On Error GoTo AUDIT_ERR

AUDITBENF = False

If glbSQL Or glbOracle Then
    SQLQ = "INSERT INTO HRAUDIT ("
    SQLQ = SQLQ & " AU_COMPNO"
    SQLQ = SQLQ & ",AU_EMPNBR"
    SQLQ = SQLQ & ",AU_PTUPL"
    SQLQ = SQLQ & ",AU_DIVUPL"
    SQLQ = SQLQ & ",AU_NEWEMP"
    SQLQ = SQLQ & ",AU_BCODE"
    SQLQ = SQLQ & ",AU_COVER"
    SQLQ = SQLQ & ",AU_MAXDOL"
    SQLQ = SQLQ & ",AU_EDATE"
    SQLQ = SQLQ & ",AU_LDATE"
    SQLQ = SQLQ & ",AU_LUSER"
    SQLQ = SQLQ & ",AU_LTIME"
    SQLQ = SQLQ & ",AU_UPLOAD"
    SQLQ = SQLQ & ",AU_TYPE"
    SQLQ = SQLQ & ",AU_TCOST"
    SQLQ = SQLQ & ",AU_PREMIUM"
    SQLQ = SQLQ & ",AU_PCE"
    SQLQ = SQLQ & ",AU_PCC"
    SQLQ = SQLQ & ",AU_PPAMT"
    SQLQ = SQLQ & ",AU_PER"
    SQLQ = SQLQ & ",AU_BAMT"
    SQLQ = SQLQ & ",AU_UNITCOST"
    SQLQ = SQLQ & ",AU_MTHECOST"
    SQLQ = SQLQ & ",AU_MTHCCOST"
    SQLQ = SQLQ & " )"
    SQLQ = SQLQ & " SELECT"
    SQLQ = SQLQ & " '001'"
    SQLQ = SQLQ & ",BF_EMPNBR"
    SQLQ = SQLQ & ",ED_PT"
    SQLQ = SQLQ & ",ED_DIV"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",BF_BCODE"
    SQLQ = SQLQ & ",BF_COVER"
    SQLQ = SQLQ & ",BF_MAXDOL"
    SQLQ = SQLQ & ",BF_EDATE"
    SQLQ = SQLQ & ",BF_LDATE"
    SQLQ = SQLQ & ",'" & glbUserID & "'"
    SQLQ = SQLQ & ",BF_LTIME"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",'" & ACTX & "'"
    SQLQ = SQLQ & ",BF_TCOST"
    SQLQ = SQLQ & ",BF_PREMIUM"
    SQLQ = SQLQ & ",BF_PCE"
    SQLQ = SQLQ & ",BF_PCC"
    SQLQ = SQLQ & ",BF_PPAMT"
    SQLQ = SQLQ & ",BF_PER"
    SQLQ = SQLQ & ",BF_AMT"
    SQLQ = SQLQ & ",BF_UNITCOST"
    SQLQ = SQLQ & ",BF_MTHECOST"
    SQLQ = SQLQ & ",BF_MTHCCOST"
    If glbOracle Then
        SQLQ = SQLQ & " FROM HRBENFT, HREMP WHERE HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND BF_LUSER='999999998' "
    Else
        SQLQ = SQLQ & " FROM HRBENFT INNER JOIN HREMP ON HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE BF_LUSER='999999998' "
    End If
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    TC.Open "SELECT * FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'Fields added by Bryan 02/Dec/05 Ticket#9899
    Dim strFields As String
    strFields = "AU_PTUPL, AU_DIVUPL, AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, "
    strFields = strFields & "AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_BCODE, AU_COVER, "
    strFields = strFields & "AU_MAXDOL, AU_EDATE, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, AU_PPAMT, AU_PER, AU_BAMT, "
    strFields = strFields & "AU_UNITCOST, AU_MTHECOST, AU_MTHCCOST, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, "
    strFields = strFields & "AU_UPLOAD, AU_TYPE "
    TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    xADD = False
    Do Until TC.EOF
        TA.AddNew
        TB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR=" & TC("BF_EMPNBR"), gdbAdoIhr001, adOpenKeyset
        If Not TB.EOF Then
            TA("AU_PTUPL") = TB("ED_PT")
            TA("AU_DIVUPL") = TB("ED_DIV")
        End If
       
        TB.Close
        TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR": TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL": TA("AU_EARN_TABL") = "EARN"
        TA("AU_NEWEMP") = "N"
        TA("AU_BCODE") = TC("BF_BCODE")
        TA("AU_COVER") = TC("BF_COVER")
        TA("AU_MAXDOL") = TC("BF_MAXDOL")
        TA("AU_EDATE") = TC("BF_EDATE")

        TA("AU_TCOST") = TC("BF_TCOST")
        TA("AU_PREMIUM") = TC("BF_PREMIUM")
        TA("AU_PCE") = TC("BF_PCE")
        TA("AU_PCC") = TC("BF_PCC")
        TA("AU_PPAMT") = TC("BF_PPAMT")
        TA("AU_PER") = TC("BF_PER")
        TA("AU_BAMT") = TC("BF_AMT")
        TA("AU_UNITCOST") = TC("BF_UNITCOST")
        TA("AU_MTHECOST") = TC("BF_MTHECOST")
        TA("AU_MTHCCOST") = TC("BF_MTHCCOST")
        
        TA("AU_COMPNO") = "001"
        TA("AU_EMPNBR") = TC("BF_EMPNBR")
        TA("AU_LDATE") = Date
        TA("AU_LUSER") = glbUserID
        TA("AU_LTIME") = Time$
        TA("AU_UPLOAD") = "N"
        TA("AU_TYPE") = ACTX
        TA.Update
        TC.MoveNext
    Loop
    TC.Close
End If
If glbWFC Then
    Call GetPayID(Date, Date)
End If
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function chkMUBENEFITS()
Dim x%, SQLQ As String, Msg As String

chkMUBENEFITS = False

On Error GoTo chkMUBENEFITS_Err
If optBenefit(1) Then
    If Len(clpBGroup(1).Text) < 1 Then
        MsgBox "New Benefit Group is a required field"
        clpBGroup(1).SetFocus
        Exit Function
    End If
    If clpBGroup(0).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpBGroup(0).SetFocus
        Exit Function
    End If
    If clpBGroup(1).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpBGroup(1).SetFocus
        Exit Function
    End If
    
    chkMUBENEFITS = True
    Exit Function
End If
Call medTCost_Change

If medEECost.Text = "" Then medEECost.Text = 0
If medCompCost.Text = "" Then medCompCost.Text = 0
If medMCCOST.Text = "" Then medMCCOST.Text = 0
If medMECOST.Text = "" Then medMECOST.Text = 0
If medTCost.Text = "" Then medTCost.Text = 0

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If
If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If
If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If
For x% = 1 To 6
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Text = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x%).SetFocus
        Exit Function
    End If
    If x% = 4 Then
        If Len(clpCode(4).Text) < 1 Then
            MsgBox "Benefit is a required field"
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
Next x%
If Len(dlpEDate.Text) > 0 Then
    If Not IsDate(dlpEDate.Text) Then
        MsgBox "Effective Date is not a valid date."
        dlpEDate.SetFocus
        Exit Function
    End If
Else
    If Actn = "A" Then
        MsgBox "Effective Date is required."
        dlpEDate.SetFocus
        Exit Function
    End If
End If
If Actn = "D" Then GoTo BpChk

'--------------------------- Add/Modify
'If Actn = "M" And Len(txtCovType) = 0 Then
'    MsgBox "Coverage is required."
'    txtCovType.SetFocus
'    Exit Function
'End If
If Len(medPayPeriodAmount) > 0 Then
   If Not IsNumeric(medPayPeriodAmount) Then
       MsgBox "Pay Period Amount is Invalid"
       medPayPeriodAmount.SetFocus
       Exit Function
   End If
Else
   medPayPeriodAmount = 0
End If

If Len(medMaxAmnt) > 0 Then
    If Not IsNumeric(medMaxAmnt) Then
        MsgBox "Maximum Amount is Invalid"
        medMaxAmnt.SetFocus
        Exit Function
    End If
Else
    medMaxAmnt = 0
End If
'--------added by Jaddy 11/2/99 begin
If comSalDepn = "Yes" Then
    If Len(medMinCover) > 0 Then
        If Not IsNumeric(medMinCover) Then
            MsgBox "Minimum Coverage Must Entry a Number ", 16
            If medMinCover.Enabled Then medMinCover.SetFocus
            Exit Function
        End If
    Else
        medMinCover = 0
    End If
    If Len(medMaxCover) > 0 Then
        If Not IsNumeric(medMaxCover) Then
            MsgBox "Maximum Coverage Must Entry a Number ", 16
            If medMaxCover.Enabled Then medMaxCover.SetFocus
            Exit Function
        Else
            If Val(medMaxCover) > 0 And Val(medMinCover) > 0 Then
                If Val(medMaxCover) < Val(medMinCover) Then
                    MsgBox "Maximum Coverage Must Be Greater Then Minimum Coverage", 16
                    If medMaxCover.Enabled Then medMaxCover.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        medMaxCover = 0
    End If
    If Len(Trim(medSalFactor)) > 0 Then
        If Not IsNumeric(medSalFactor) Then
            MsgBox "Salary Factor Must Entry a Number ", 16
            If medSalFactor.Enabled Then medSalFactor.SetFocus
            Exit Function
'        Else
'            If Val(medSalFactor) = 0 Then
'                MsgBox "Salary Factor Must Be Greater Then 0", 16
'                If medSalFactor.Enabled Then medSalFactor.SetFocus
'                Exit Function
'            End If
        End If
    Else
        medSalFactor = 0
'          MsgBox "Salary Factor Must Be Greater Then 0", 16
'          If medSalFactor.Enabled Then medSalFactor.SetFocus
'          Exit Function
    End If

Else
'--------added by Jaddy 11/2/99 end
    If Len(medCovAmount) > 0 Then
        If Not IsNumeric(medCovAmount) Then
            MsgBox "Coverage Amount is Invalid", 48
            medCovAmount.SetFocus
            Exit Function
        End If
        If optActual(1) And medCovAmount = 0 Then
            MsgBox "Coverage Amount is Required", 48
            medCovAmount.SetFocus
            Exit Function
        End If
    Else
        If optActual(1) Then
            MsgBox "Coverage Amount is Required", 48
            medCovAmount.SetFocus
            Exit Function
        Else
            medCovAmount = 0
        End If
        
    End If
End If 'jaddy 11/2/99
If Len(medUnitCost) > 0 Then
    If Not IsNumeric(medUnitCost) Then
        MsgBox "Per Unit is Invalid.", 48
        medUnitCost.SetFocus
        Exit Function
    End If
    If optActual(1) And medUnitCost = 0 Then
        MsgBox "Per Unit is required."
        medUnitCost.SetFocus
        Exit Function
    End If
Else
    If optActual(1) Then
        MsgBox "Per Unit is required."
        medUnitCost.SetFocus
        Exit Function
    End If
    medUnitCost = 0
End If
If Len(txtPer) > 0 Then
    If Not IsNumeric(txtPer) Then
        MsgBox "Per Unit Cost is Invalid.", 48
        txtPer.SetFocus
        Exit Function
    End If
    If optActual(1) And txtPer = 0 Then
        MsgBox "Per Unit Cost is required."
        txtPer.SetFocus
        Exit Function
    End If
Else
    If optActual(1) Then
        MsgBox "Per Unit Cost is required."
        txtPer.SetFocus
        Exit Function
    End If
    txtPer = 0
End If

If Len(medPPComp) <= 0 Then
    MsgBox "Company Percentage Paid is required"
    medPPComp.SetFocus
    Exit Function
End If
If Not IsNumeric(medPPComp) Then
    MsgBox "Company Percentage Paid is Invalid", 48
    medPPComp.SetFocus
    Exit Function
End If
If medPPComp > 1 Or medPPComp < 0 Then
    MsgBox "Company Percentage Paid is Invalid", 48
    medPPComp.SetFocus
    Exit Function
End If
'If Len(medMCCOST) > 0 Then      'laura 02/27/98
'  If Not IsNumeric(medMCCOST) Then
'      MsgBox "Monthly Company Cost paid is Invalid", 48
'      medMCCOST.SetFocus
'      Exit Function
'  End If
'Else
'   If comSalDepn <> "Yes" Then medMCCOST = 0      'jaddy 11/3/99
'End If
'If Len(medMECOST) > 0 Then      'laura 02/27/98
'  If Not IsNumeric(medMECOST) Then
'      MsgBox "Monthly Employee Cost paid is Invalid", 48
'      medMECOST.SetFocus
'      Exit Function
'  End If
'Else
'  If comSalDepn <> "Yes" Then medMECOST = 0     'jaddy 11/3/99
'End If
'------Jaddy 11/3/99 changed begin
If optActual(0) Then
    If Len(medTCost) > 0 Then
        If Not IsNumeric(medTCost) Then
            MsgBox "Total Cost is Invalid", 48
            medTCost.SetFocus
            Exit Function
        Else
'            If medTCost = 0 Then
'                MsgBox "Total Cost is required"
'                medTCost.SetFocus
'                Exit Function
'            End If
        End If
    Else
        medTCost = 0
'        MsgBox "Total Cost is required"
'        medTCost.SetFocus
'        Exit Function
    End If
End If
If Not elpEEID.ListChecker Then
    Exit Function
End If

'------Jaddy 11/3/99 changed end
BpChk:
chkMUBENEFITS = True

Exit Function

chkMUBENEFITS_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEbenefit", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Public Sub cmdClose_Click()
Unload Me

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%, EmailResponse%
Dim recCount As Integer

strEMPLIST = ""
strTermEMPLIST = ""

If Not gSec_Upd_Benefits Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "D"
fglbDelete% = True
fglbAdd% = False
fglbModify% = False

If Not chkMUBENEFITS() Then Exit Sub

Title$ = "Mass Benefits Records Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Benefit Master Record " Else Msg$ = Msg$ & " Benefit Master Records "
    Msg$ = Msg$ & " to Delete. " & vbCrLf & vbCrLf & " Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Benefit Master records found to delete."
    GoTo End_Note
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

'Release 8.1 - Send Email Notification but prompt the user first
flgSendEmail = False
If gsEMAIL_ONBENEFIT Then
    Msg = "Do you want to send Email Notification automatically for each employees affected by this delete?"
    EmailResponse% = MsgBox(Msg, vbYesNo + vbQuestion, "Benefit Delete Email Notification")
    If EmailResponse% = vbNo Then
        flgSendEmail = False
    Else
        flgSendEmail = True
    End If
End If
    
'Release 8.1 - Email Notification - only if Benefit got updated
glbBenDeleted = "False"

x% = modDelRecs()

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
     'Release 8.1 - Benefit Delete - Send Email Notification for each as per Jerry's requirement.
    If flgSendEmail Then
        'If glbBenDeleted = "True" Then
            Call EmailNotification("DELETE")
        'End If
    End If
       
    MsgBox Str(XUpdCount) & " Records Deleted Successfully"
    If Response% = IDYES Then    ' Evaluate response
        Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        'report name
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Benefits Master - Employee Details'"
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        End If
        Me.vbxCrystal.Connect = RptODBC_SQL
        'If glbSQL Or glbOracle Then
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        'Else
        '    Me.vbxCrystal.Connect = "PWD=petman;"
        '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
        'End If
        
        ' window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        'MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        'MDIMain.Timer1.Enabled = True
        'Call set_PrintState(True)
    
    End If
Else
    MsgBox "No Records Deleted!"
End If

End_Note:

Exit Sub


Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "ATTEND", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%, EmailResponse%
Dim recCount As Integer

On Error GoTo Mod_Err

strEMPLIST = ""
strTermEMPLIST = ""

If Not gSec_Upd_Benefits Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "M"
fglbDelete% = False
fglbAdd% = False
fglbModify% = True

If Not chkMUBENEFITS() Then Exit Sub

Title$ = "Mass Update Benefits"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to update all Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

'Get total record counts to update
If optBenefit(1) Then
    recCount = getRecordCount_ModifyBenGrp
    Msg$ = Str(recCount)
    
    If recCount > 0 Then
        If recCount = 1 Then Msg$ = Msg$ & " Benefit Group Record " Else Msg$ = Msg$ & " Benefit Group Records "
    Else
        MsgBox "No Benefit Group records found to update."
        GoTo End_Note
    End If
Else
    recCount = getRecordCount_ModifyBenMst
    Msg$ = Str(recCount)
    
    If recCount > 0 Then
        If recCount = 1 Then Msg$ = Msg$ & " Benefit Master Record " Else Msg$ = Msg$ & " Benefit Master Records "
    Else
        MsgBox "No Benefit Master records found to update."
        GoTo End_Note
    End If
End If
Msg$ = Msg$ & "to Update. " & vbCrLf & vbCrLf & "Do you want to proceed?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response = IDNO Then
    Exit Sub
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If optBenefit(1) Then
    If Not modUpdGroup() Then Exit Sub
Else
    'Release 8.1 - Send Email Notification but prompt the user first
    flgSendEmail = False
    If gsEMAIL_ONBENEFIT Then
        Msg = "Do you want to send Email Notification automatically for each employees affected by this update?"
        EmailResponse% = MsgBox(Msg, vbYesNo + vbQuestion, "Benefit Update Email Notification")
        If EmailResponse% = vbNo Then
            flgSendEmail = False
        Else
            flgSendEmail = True
        End If
    End If
        
    'Release 8.1 - Email Notification - only if Benefit got updated
    glbBenChanged = "False"
        
    If Not modUpdRecs() Then Exit Sub
    
    'Release 8.1 - Send Email Notifications for employees updated only
    If XUpdCount > 0 Then
         'Release 8.1 - Employee Updated - Send Email Notification
        If flgSendEmail Then
            'If glbBenChanged <> "False" Then
                Call EmailNotification("UPDATE")
            'End If
        End If
    End If
End If


Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Updated Successfully"
    
    If Response% = IDYES Then    ' Evaluate response
        Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        ' report name
      
        If chkEmployee(0) Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
        ElseIf chkEmployee(1) Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZTEmpList.rpt"
        End If
    
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Benefits Master - Employee Details'"
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        End If
        Me.vbxCrystal.Connect = RptODBC_SQL
        'If glbSQL Or glbOracle Then
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        'Else
        '    Me.vbxCrystal.Connect = "PWD=petman;"
        '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
        'End If
        
        ' window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        'MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        'MDIMain.Timer1.Enabled = True
        'Call set_PrintState(True)
    
    End If
Else
    MsgBox "Employees for this selection do not exist!"
End If

End_Note:

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%, EmailResponse%
Dim recCount  As Integer

On Error GoTo AddN_Err

strEMPLIST = ""
strTermEMPLIST = ""

If Not gSec_Upd_Benefits Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "A"
fglbAdd% = True
fglbDelete% = True
fglbModify% = False

If Not chkMUBENEFITS() Then Exit Sub

Title$ = "Mass Records Benefits"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Benefit Master Record " Else Msg$ = Msg$ & " Benefit Master Records "
    Msg$ = Msg$ & "to Add. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No employee found to add Benefit record."
    GoTo End_Note
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.


'Release 8.1 - Send Email Notification but prompt the user first
flgSendEmail = False
If gsEMAIL_ONBENEFIT Then
    Msg = "Do you want to send Email Notification automatically for each employees affected by this add?"
    EmailResponse% = MsgBox(Msg, vbYesNo + vbQuestion, "Benefit Add Email Notification")
    If EmailResponse% = vbNo Then
        flgSendEmail = False
    Else
        flgSendEmail = True
    End If
End If
    
'Release 8.1 - Email Notification - only if Benefit got updated
glbBenAdded = "False"

If Not modInsRecs() Then Exit Sub

Screen.MousePointer = DEFAULT

If XUpdCount > 0 Then
'    If gsEMAIL_ONBENEFIT Then
'         If Len(MailBody) > 0 Then
'            If XUpdCount = 1 Then
'                xStr = "The following employee has "
'            Else
'                xStr = "The following employees have "
'            End If
'            xStr = xStr & " new benefit." & vbCrLf
'            xStr = xStr & "Benefit Name: " & GetTABLDesc("BNCD", clpCode(4)) & vbCrLf
'            xStr = xStr & "Effective Date: " & dlpEDate & vbCrLf & vbCrLf
'            MailBody = xStr & MailBody
'            Screen.MousePointer = DEFAULT
'            Call imgEmail_Click
'         End If
'    End If

     'Release 8.1 - Benefit Added - Send Email Notification for each as per Jerry's requirement therefore commenting the above.
    If flgSendEmail Then
        'If glbBenAdded = "True" Then
            Call EmailNotification("ADD")
        'End If
    End If

    MsgBox Str(XUpdCount) & " Records Added Successfully"
    If Response% = IDYES Then    ' Evaluate response
        Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        ' report name
      
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Benefits Master - Employee Details'"
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        End If
        Me.vbxCrystal.Connect = RptODBC_SQL
        'If glbSQL Or glbOracle Then
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        'Else
        '    Me.vbxCrystal.Connect = "PWD=petman;"
        '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
        'End If
        
        ' window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        'MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        'MDIMain.Timer1.Enabled = True
        'Call set_PrintState(True)
    
    End If
Else
    MsgBox "No Records Added."
End If

End_Note:

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Function GetEmpName(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim xStr, SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xStr = "Employee #:" & (xEmpNo) & " Name: " & rsTemp("ED_FNAME") & " " & rsTemp("ED_SURNAME")
    End If
    rsTemp.Close
    GetEmpName = xStr
End Function

Public Sub imgEmail_Click()
Dim xEmail
On Error GoTo Email_Err
    If gsEMAIL_ONBENEFIT Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = xEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice"
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        Else
            MsgBox "There is no email for Email Notification on Benefits on Company Preference screen. "
        End If


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

Private Sub comSalDepn_Change()
comSalDepn_Click
End Sub

Private Sub comSalDepn_Click()
If comSalDepn = "Yes" Then
    comRndFactor.Enabled = True
    medMinCover.Enabled = True
    medMaxCover.Enabled = True
    medSalFactor.Enabled = True
    medCovAmount.Enabled = False
    Set_SalCover
Else
    comRndFactor.Enabled = False
    medMinCover.Enabled = False
    medMaxCover.Enabled = False
    medSalFactor.Enabled = False
    medCovAmount.Enabled = True
    comRndFactor = ""
    medMinCover = ""
    medMaxCover = ""
    medSalFactor = ""
    medCovAmount = 0
End If
End Sub

Private Sub comSalDepn_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Set_SalCover()
'added by jaddy 9/8/99
''Hidden by Jaddy 11/2/99
'Dim xSalFactor, xRndFactor, xMaxCover, xMinCover, xCovAmount
'Dim xSalary
'xSalary = CrtSalary(glbLEE_ID)
'xSalFactor = Val(medSalFactor)
'xRndFactor = Val(comRndFactor)
'xMaxCover = Val(medMaxCover)
'xMinCover = Val(medMinCover)
'xCovAmount = xSalary * xSalFactor
'If xMaxCover > xMinCover Then
'    If xCovAmount > xMaxCover Then xCovAmount = xMaxCover
'    If xCovAmount < xMinCover Then xCovAmount = xMinCover
'End If
'If xRndFactor <> 0 Then xCovAmount = Round(xCovAmount / xRndFactor) * xRndFactor
'medCovAmount = xCovAmount
medCovAmount = "From System"
End Sub

Private Sub comRndFactor_Change()
Call Set_SalCover
End Sub

Private Sub comRndFactor_Click()
Call Set_SalCover
End Sub

Private Sub comRndFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSalFactor_Change()
Call Set_SalCover
End Sub

Private Sub medSalFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMinCover_Change()
Call Set_SalCover
End Sub

Private Sub medMinCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMaxCover_Change()
Call Set_SalCover
End Sub

Private Sub medMaxCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
Private Sub Form_Load()

glbOnTop = "FRMUBENEFITS"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

lblAP = "A"
Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)
If glbCompSerial = "S/N - 2227W" Then
    clpCode(5).MaxLength = 6
End If

If glbLinamar Then
    clpCode(4).MaxLength = 8
    clpCode(5).MaxLength = 8
    optBenefit(1).Enabled = False  'It is not available for Linamar
End If

If glbCompSerial = SN_CASEY Then
    medPPComp.Enabled = False
    medTCost.Enabled = False
    medMECOST.Enabled = True
    medMCCOST.Enabled = True
End If

If Not glbSQL Then
    chkEmployee(1).Enabled = False
End If

Call INI_Controls(Me)

medPPComp = 0

Call optActual_Click(0, 1)

If glbMulti Then textMulti.Visible = True
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmUBENEFITS = Nothing  'carmen apr 2000

End Sub

Private Sub lblAP_Click()
If lblAP = "A" Then
    optActual(0).Value = True
Else
    optActual(1).Value = True
End If

End Sub

Private Sub medCompCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCovAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCovAmount_Change()
If optActual(1).Value = True Then
    If comSalDepn = "Yes" Then
        medTCost = "From System"
    Else
        If Not IsNumeric(medCovAmount) Then medCovAmount = 0
        If Not IsNumeric(txtPer) Then txtPer = 0
        If Not IsNumeric(medUnitCost) Then medUnitCost = 0
        If txtPer > 0 And medUnitCost > 0 Then
            medTCost = medCovAmount / txtPer * medUnitCost
        Else
            medTCost = 0
        End If
    End If
End If
End Sub

Private Sub medEECost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMaxAmnt_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMCCOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMECOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPayPeriodAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPPComp_GotFocus()
Call SetPanHelp(ActiveControl)
If Not IsNumeric(medPPComp) Then
    medPPComp = 0
Else
    medPPComp = medPPComp * 100
End If
End Sub

Private Sub medPPComp_Change()
    If glbCompSerial <> SN_CASEY Then
        medPPE = 1 - Val(medPPComp) / 100
        Call medTCost_Change
    End If
End Sub

Private Sub medPPComp_LostFocus()
If Len(medPPComp) > 0 Then
    If IsNumeric(medPPComp) Then
        medCompCost = Val(medTCost) * Val(medPPComp) / 100 '1
        medPPComp = Val(medPPComp) / 100 '2
        medPPE = 1 - medPPComp
        Call medTCost_Change
        medCompCost.Visible = True
    End If
End If
End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTCost_Change()
'changed by jaddy 9/10/99---------------------------------
' This is Called by :
' medPPComp_LostFocus & medTCost_LostFocus & chkMUBENEFITS
'---------------------------------------------------------
If medTCost = "From System" Then
    medEECost = "From System"
    medCompCost = "From System"
    medMECOST = "From System"
    medMCCOST = "From System"
Else
    ' dkostka - 01/08/2001 - CalcBeneCasey handles calculations for Casey House, don't do them here
    '   or we get into a loop.
    If glbCompSerial <> SN_CASEY Then
        medEECost = Val(medTCost) * Val(medPPE)
        medCompCost = Val(medTCost) * Val(medPPComp)
        medMECOST = Round(Val(medEECost) / 12, 2)
        medMCCOST = Round(Val(medCompCost) / 12, 2)
    End If
End If
End Sub

Private Sub medTCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUnitCost_Change()
Call medCovAmount_Change
End Sub

Private Function modDelRecs()
Dim RecsAffected As Long
Dim SQLQ As String

modDelRecs = False
On Error GoTo modDelRecs_Err
Screen.MousePointer = HOURGLASS

Call getWSQLQ("D")

If chkEmployee(0) Then
    
    SQLQ = "UPDATE HRBENFT SET BF_LUSER = '999999998' "
    SQLQ = SQLQ & " WHERE " & fglbWSQLQ
    SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & " )"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans

    If Not AUDITBENF("D") Then MsgBox "ERROR - AUDIT FILE"
    
    SQLQ = "SELECT BF_EMPNBR FROM HRBENFT WHERE BF_LUSER='999999998' "
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do While Not RSEMPLIST.EOF
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & RSEMPLIST("BF_EMPNBR")
        Else
            strEMPLIST = strEMPLIST & RSEMPLIST("BF_EMPNBR")
        End If
        RSEMPLIST.MoveNext
    Loop
    RSEMPLIST.Close
    
    SQLQ = "DELETE FROM HRBENFT WHERE BF_LUSER='999999998' "
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ, RecsAffected
    XUpdCount = RecsAffected
    gdbAdoIhr001.CommitTrans
End If

If chkEmployee(1) Then
    SQLQ = "SELECT BF_EMPNBR FROM Term_HRBENFT"
    SQLQ = SQLQ & " WHERE " & fglbWSQLQ
    SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & " )"
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do While Not RSEMPLIST.EOF
        If Len(strTermEMPLIST) > 0 Then
            strTermEMPLIST = strTermEMPLIST & "," & RSEMPLIST("BF_EMPNBR")
        Else
            strTermEMPLIST = strTermEMPLIST & RSEMPLIST("BF_EMPNBR")
        End If
        RSEMPLIST.MoveNext
    Loop
    RSEMPLIST.Close
    
    SQLQ = "DELETE FROM Term_HRBENFT"
    SQLQ = SQLQ & " WHERE " & fglbWSQLQ
    SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & " )"
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute SQLQ, RecsAffected
    XUpdCount = XUpdCount + RecsAffected
    gdbAdoIhr001X.CommitTrans
End If


modDelRecs = True

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteAttend", "Delete")
modDelRecs = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()
Dim RecsAffected As Long
Dim SQLQ As String, USQLQ As String
Dim SalaryStr
Dim x%
Dim TCostStr
Dim CostStr
Dim strDays As String
Dim rsBenAdd As New ADODB.Recordset
Dim SQLQEmail As String

modInsRecs = False

On Error GoTo modInsRecs_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ("A")

If GetLeapYear(Year(Date)) Then
    strDays = "366"
Else
    strDays = "365"
End If

SQLQ = SQLQ & "INSERT INTO HRBENFT "
SQLQ = SQLQ & "(BF_EMPNBR, BF_BCODE,  BF_EDATE, BF_COVER,  BF_SALARYDEPENDANT, BF_UNITCOST,"
If comSalDepn = "Yes" Then SQLQ = SQLQ & "BF_MAXIMUM,  BF_MINIMUM, BF_FACTOR, BF_ROUND, BF_NEXTNEAREST, "
SQLQ = SQLQ & "BF_AMT, BF_PPAMT, BF_PCE, BF_PCC, BF_PREMIUM, BF_PER, "
SQLQ = SQLQ & " BF_TCOST, BF_ECOST, BF_MTHECOST, BF_CCOST, BF_MTHCCOST, BF_DWM,"
If txtTAXBEN <> "" Then SQLQ = SQLQ & "BF_TAXBEN, "
' danielk - 02/05/2003 - added pre/post tax
If txtPreAftTax <> "" Then SQLQ = SQLQ & "BF_PTAX, "
' danielk - 02/05/2003 - end
SQLQ = SQLQ & "BF_LDATE, BF_LTIME, BF_LUSER) "

SQLQ = SQLQ & "SELECT HREMP.ED_EMPNBR AS BF_EMPNBR, "
If glbLinamar Then
    SQLQ = SQLQ & "'" & IIf(clpDiv.Text = "", "ALL", clpDiv.Text) & clpCode(4).Text & "' as BF_BCODE, "
Else
    SQLQ = SQLQ & "'" & clpCode(4).Text & "' as BF_BCODE, "
End If

SQLQ = SQLQ & Date_SQL(dlpEDate.Text) & " as BF_EDATE, "

SQLQ = SQLQ & "'" & txtCovType & "' as BF_COVER, "
SQLQ = SQLQ & "'" & Left(comSalDepn, 1) & "' as BF_SALARYDEPENDANT, "
SQLQ = SQLQ & medUnitCost & " as BF_UNITCOST, "   'laura Dec 12, 1997
If glbSQL Or glbOracle Then
    SalaryStr = "(CASE WHEN HR_SALARY_HISTORY.SH_SALCD='A' THEN 1 WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN 12 " & _
                "WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN " & strDays & " ELSE HR_SALARY_HISTORY.SH_WHRS*52 END) "
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
        SalaryStr = SalaryStr & "* (CASE WHEN HR_SALARY_HISTORY.SH_NFAC_SALARY<>0 AND HR_SALARY_HISTORY.SH_NFAC_SALARY IS NOT NULL THEN HR_SALARY_HISTORY.SH_NFAC_SALARY ELSE HR_SALARY_HISTORY.SH_SALARY END) * " & Val(medSalFactor)
    Else
        SalaryStr = SalaryStr & "* HR_SALARY_HISTORY.SH_SALARY * " & Val(medSalFactor)
    End If
    SalaryStr = "(CASE WHEN " & Val(medMinCover) & "> 0 And " & Val(medMinCover) & " >" & SalaryStr & " THEN " & Val(medMinCover) & " ELSE " & SalaryStr & " END )"
    SalaryStr = "(CASE WHEN " & Val(medMaxCover) & "> 0 And " & Val(medMaxCover) & " <" & SalaryStr & " THEN " & Val(medMaxCover) & " ELSE " & SalaryStr & " END )"
    If Val(comRndFactor) > 0 Then SalaryStr = "ROUND(" & SalaryStr & "/" & Val(comRndFactor) & _
        IIf(optRound(0), "+0", "+.49") & ",0) *" & Val(comRndFactor)
Else
    SalaryStr = "iif(HR_SALARY_HISTORY.SH_SALCD='M',HR_SALARY_HISTORY.SH_SALARY * 12, " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52), " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='D',HR_SALARY_HISTORY.SH_SALARY * " & strDays & ", " & _
    "HR_SALARY_HISTORY.SH_SALARY))) * " & Val(medSalFactor)
    SalaryStr = "IIf(" & Val(medMinCover) & "> 0 And " & Val(medMinCover) & " >" & SalaryStr & "," & Val(medMinCover) & "," & SalaryStr & ")"
    SalaryStr = "IIf(" & Val(medMaxCover) & "> 0 And " & Val(medMaxCover) & " <" & SalaryStr & "," & Val(medMaxCover) & "," & SalaryStr & ")"
    If Val(comRndFactor) > 0 Then SalaryStr = "ROUND(" & SalaryStr & "/" & Val(comRndFactor) & _
        IIf(optRound(0), "+0", "+.5") & ") *" & Val(comRndFactor)
End If
If comSalDepn = "Yes" Then
    SQLQ = SQLQ & Val(medMaxCover) & " as BF_MAXIMUM, "
    SQLQ = SQLQ & Val(medMinCover) & " as BF_MINIMUM, "
    SQLQ = SQLQ & Val(medSalFactor) & " as BF_FACTOR, "
    SQLQ = SQLQ & Val(comRndFactor) & " as BF_ROUND, "
    SQLQ = SQLQ & "'" & IIf(optRound(0), "R", "N") & "' as BF_NEXTNEAREST, "
    SQLQ = SQLQ & SalaryStr & " as BF_AMT,  "
Else
    SQLQ = SQLQ & medCovAmount & " as BF_AMT,  "
End If
SQLQ = SQLQ & medPayPeriodAmount & " as BF_PPAMT, "
SQLQ = SQLQ & medPPE & " as BF_PCE, "
SQLQ = SQLQ & medPPComp & " as BF_PCC, "
SQLQ = SQLQ & "'" & lblAP & "' as BF_PREMIUM, "
SQLQ = SQLQ & txtPer & " as BF_PER, "
If comSalDepn = "Yes" And optActual(1) Then
    TCostStr = "(" & SalaryStr & ")/" & txtPer & "*" & medUnitCost
        SQLQ = SQLQ & TCostStr & " as BF_TCOST, "
    CostStr = TCostStr & "*" & medPPE
        SQLQ = SQLQ & CostStr & " as BF_ECOST, "
    CostStr = CostStr & "/12"
        SQLQ = SQLQ & CostStr & " as BF_MTHECOST, "
    CostStr = TCostStr & "*" & medPPComp
        SQLQ = SQLQ & CostStr & " as BF_CCOST, "
    CostStr = CostStr & "/12"
        SQLQ = SQLQ & CostStr & " as BF_MTHCCOST, "
Else
    SQLQ = SQLQ & medTCost & " as BF_TCOST, "
    SQLQ = SQLQ & medEECost & " as BF_ECOST, "
    SQLQ = SQLQ & medMECOST & " as BF_MTHECOST, "    'laura 02/24/98
    SQLQ = SQLQ & medCompCost & " as BF_CCOST, "
    SQLQ = SQLQ & medMCCOST & " as BF_MTHCCOST, "    'laura 02/24/98
End If
'Jaddy Changed
If glbCElgin Then
    SQLQ = SQLQ & "'D' as BF_DWM, "
Else
    SQLQ = SQLQ & "NULL as BF_DWM, "
End If
If txtTAXBEN <> "" Then SQLQ = SQLQ & "'" & txtTAXBEN & "' as BF_TAXBEN, "     'laura 02/24/98
' danielk - 02/05/2003 - Added pre/post tax
If txtPreAftTax.Text <> "" Then SQLQ = SQLQ & "'" & txtPreAftTax.Text & "' AS BF_PTAX, "
' danielk - 02/05/2003 - end
SQLQ = SQLQ & Date_SQL(Date) & " as BF_LDATE, "
SQLQ = SQLQ & "'" & Time$ & "' as BF_LTIME, "
SQLQ = SQLQ & "'999999998'" & " as BF_LUSER "
SQLQ = SQLQ & " FROM HREMP "

If comSalDepn = "Yes" Then
    If glbOracle Then
        SQLQ = SQLQ & " ,HR_SALARY_HISTORY WHERE HREMP.ED_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR "
    Else
        SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HREMP.ED_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR "
    End If
End If
If InStr(SQLQ, "WHERE") <> 0 Then
    SQLQ = SQLQ & " AND " & fglbESQLQ
Else
    SQLQ = SQLQ & " WHERE " & fglbESQLQ
End If
If comSalDepn = "Yes" Then SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 "
XUpdCount = 0
If chkEmployee(0) Then
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    If Not AUDITBENF("A") Then MsgBox "ERROR - AUDIT FILE"
    
    'Release 8.1 - Email Notification on New Benefit added
    'Get the list of Employee # who got new Benefit added
    '----------------------------------------------------------------------------
    SQLQEmail = "SELECT BF_EMPNBR FROM HRBENFT WHERE BF_LUSER='999999998' "
    EmpListAdd = ""
    rsBenAdd.Open SQLQEmail, gdbAdoIhr001, adOpenStatic
    Do While Not rsBenAdd.EOF
        If Len(EmpListAdd) > 0 Then
            EmpListAdd = EmpListAdd & "," & rsBenAdd("BF_EMPNBR")
        Else
            EmpListAdd = rsBenAdd("BF_EMPNBR")
        End If
        rsBenAdd.MoveNext
    Loop
    rsBenAdd.Close
    Set rsBenAdd = Nothing
    '----------------------------------------------------------------------------
    
    USQLQ = "UPDATE HRBENFT SET BF_LUSER='" & glbUserID & "' WHERE BF_LUSER='999999998' "
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute USQLQ, RecsAffected
    XUpdCount = RecsAffected
    gdbAdoIhr001.CommitTrans
End If

If chkEmployee(1) Then
    SQLQ = Replace(SQLQ, "HREMP", "Term_HREMP")
    SQLQ = Replace(SQLQ, "HRBENFT", "Term_HRBENFT")
    SQLQ = Replace(SQLQ, "HR_SALARY_HISTORY", "Term_SALARY_HISTORY")
    SQLQ = Replace(SQLQ, "999999998", glbUserID)
    SQLQ = Replace(SQLQ, "BF_EMPNBR,", "BF_EMPNBR,TERM_SEQ,")
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute SQLQ, RecsAffected
    XUpdCount = XUpdCount + RecsAffected
    gdbAdoIhr001X.CommitTrans
End If

Dim rsEMP As New ADODB.Recordset
Dim xEmpnbr
SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & " "
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
MailBody = ""
Do While Not rsEMP.EOF
    xEmpnbr = rsEMP("ED_EMPNBR")
    Call updBenefitForSalDEPN(xEmpnbr)
    If glbGP Then 'Ticket #30111 Franks 06/13/2017
        Call Employee_GP_NewBenefitDeduction_Integration(xEmpnbr)
    End If
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & rsEMP("ED_EMPNBR")
    Else
        strEMPLIST = rsEMP("ED_EMPNBR")
    End If
    If gsEMAIL_ONBENEFIT Then
        MailBody = MailBody & GetEmpName(rsEMP("ED_EMPNBR")) & vbCrLf
    End If
    rsEMP.MoveNext
Loop


modInsRecs = True

Exit Function

modInsRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modInsRecs", "Benefits", "Insert")
modInsRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modUpdRecs()
Dim RecsAffected As Long
Dim SQLQ As String, USQLQ As String
Dim x%
Dim SalaryStr, CostStr, TCostStr
Dim strDays As String

modUpdRecs = False

On Error GoTo modUpdRecs2_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ("U")

If GetLeapYear(Year(Date)) Then
    strDays = 366
Else
    strDays = 365
End If

If glbOracle And comSalDepn = "Yes" Then
    SQLQ = "UPDATE  HRBENFT SET "
    SQLQ = SQLQ & "(BF_EMPNBR, BF_BCODE,  BF_EDATE, BF_COVER,  BF_SALARYDEPENDANT, BF_UNITCOST,"
    SQLQ = SQLQ & "BF_MAXIMUM,  BF_MINIMUM, BF_FACTOR, BF_ROUND, BF_NEXTNEAREST, "
    SQLQ = SQLQ & "BF_AMT, BF_PPAMT, BF_PCE, BF_PCC, BF_PREMIUM, BF_PER, "
    SQLQ = SQLQ & " BF_TCOST, BF_ECOST, BF_MTHECOST, BF_CCOST, BF_MTHCCOST, "
    If txtTAXBEN <> "" Then SQLQ = SQLQ & "BF_TAXBEN, "
    SQLQ = SQLQ & "BF_PTAX,BF_LDATE, BF_LTIME, BF_LUSER) ="
    
    SQLQ = SQLQ & "(SELECT HR_SALARY_HISTORY.SH_EMPNBR AS BF_EMPNBR, "
    SQLQ = SQLQ & "'" & clpCode(4).Text & "' as BF_BCODE, "
    SQLQ = SQLQ & Date_SQL(dlpEDate.Text) & " as BF_EDATE, "
    SQLQ = SQLQ & "'" & txtCovType & "' as BF_COVER, "
    SQLQ = SQLQ & "'" & Left(comSalDepn, 1) & "' as BF_SALARYDEPENDANT, "
    SQLQ = SQLQ & medUnitCost & " as BF_UNITCOST, "
    
    SalaryStr = "(CASE WHEN HR_SALARY_HISTORY.SH_SALCD='A' THEN 1 ELSE HR_SALARY_HISTORY.SH_WHRS*52 END)* HR_SALARY_HISTORY.SH_SALARY * " & Val(medSalFactor)
    SalaryStr = "(CASE WHEN " & Val(medMinCover) & "> 0 And " & Val(medMinCover) & " >" & SalaryStr & " THEN " & Val(medMinCover) & " ELSE " & SalaryStr & " END )"
    SalaryStr = "(CASE WHEN " & Val(medMaxCover) & "> 0 And " & Val(medMaxCover) & " <" & SalaryStr & " THEN " & Val(medMaxCover) & " ELSE " & SalaryStr & " END )"
    If Val(comRndFactor) > 0 Then SalaryStr = "ROUND(" & SalaryStr & "/" & Val(comRndFactor) & _
        IIf(optRound(0), "+0", "+.5") & ",0) *" & Val(comRndFactor)
    SQLQ = SQLQ & Val(medMaxCover) & " as BF_MAXIMUM, "
    SQLQ = SQLQ & Val(medMinCover) & " as BF_MINIMUM, "
    SQLQ = SQLQ & Val(medSalFactor) & " as BF_FACTOR, "
    SQLQ = SQLQ & Val(comRndFactor) & " as BF_ROUND, "
    SQLQ = SQLQ & "'" & IIf(optRound(0), "R", "N") & "' as BF_NEXTNEAREST, "
    SQLQ = SQLQ & SalaryStr & " as BF_AMT,  "
    SQLQ = SQLQ & medPayPeriodAmount & " as BF_PPAMT, "
    SQLQ = SQLQ & medPPE & " as BF_PCE, "
    SQLQ = SQLQ & medPPComp & " as BF_PCC, "
    SQLQ = SQLQ & "'" & lblAP & "' as BF_PREMIUM, "
    SQLQ = SQLQ & txtPer & " as BF_PER, "
    If optActual(1) Then
        TCostStr = "(" & SalaryStr & ")/" & txtPer & "*" & medUnitCost
            SQLQ = SQLQ & TCostStr & " as BF_TCOST, "
        CostStr = TCostStr & "*" & medPPE
            SQLQ = SQLQ & CostStr & " as BF_ECOST, "
        CostStr = CostStr & "/12"
            SQLQ = SQLQ & CostStr & " as BF_MTHECOST, "
        CostStr = TCostStr & "*" & medPPComp
            SQLQ = SQLQ & CostStr & " as BF_CCOST, "
        CostStr = CostStr & "/12"
            SQLQ = SQLQ & CostStr & " as BF_MTHCCOST, "
    Else
        SQLQ = SQLQ & medTCost & " as BF_TCOST, "
        SQLQ = SQLQ & medEECost & " as BF_ECOST, "
        SQLQ = SQLQ & medMECOST & " as BF_MTHECOST, "
        SQLQ = SQLQ & medCompCost & " as BF_CCOST, "
        SQLQ = SQLQ & medMCCOST & " as BF_MTHCCOST, "
    End If
    If txtTAXBEN <> "" Then SQLQ = SQLQ & "'" & txtTAXBEN & "' as BF_TAXBEN, "
    ' danielk - 02/05/2003 - Added pre/post tax
    SQLQ = SQLQ & "'" & txtPreAftTax & "' as BF_PTAX, "
    ' danielk - 02/05/2003 - end
    SQLQ = SQLQ & Date_SQL(Date) & " as BF_LDATE, "
    SQLQ = SQLQ & "'" & Time$ & "' as BF_LTIME, "
    SQLQ = SQLQ & "'999999998'" & " as BF_LUSER "
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_EMPNBR=HRBENFT.BF_EMPNBR "
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 )"
    SQLQ = SQLQ & " WHERE " & fglbWSQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
Else
    SQLQ = "UPDATE  HRBENFT "
    If Not glbSQL And Not glbOracle And comSalDepn = "Yes" Then
        SQLQ = SQLQ & " LEFT JOIN HR_SALARY_HISTORY ON HRBENFT.BF_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR "
    End If
    
    SQLQ = SQLQ & "SET "
    If IsDate(dlpEDate.Text) Then SQLQ = SQLQ & "BF_EDATE = " & Date_SQL(dlpEDate.Text) & ", "
    
    'Ticket #24220 - Change Cost Only
    If Not optBenefit(2) Then
        SQLQ = SQLQ & "BF_MAXDOL = " & medMaxAmnt & ", "
        SQLQ = SQLQ & "BF_SALARYDEPENDANT = '" & Left(comSalDepn, 1) & "', "
        SQLQ = SQLQ & "BF_MAXIMUM = " & Val(medMaxCover) & ", "
        SQLQ = SQLQ & "BF_MINIMUM = " & Val(medMinCover) & ", "
        SQLQ = SQLQ & "BF_FACTOR = " & Val(medSalFactor) & ", "
        SQLQ = SQLQ & "BF_ROUND = " & Val(comRndFactor) & ", "
        SQLQ = SQLQ & "BF_NEXTNEAREST= '" & IIf(optRound(0), "R", "N") & "', "
    Else
        If IsNumeric(medMaxAmnt) Then SQLQ = SQLQ & "BF_MAXDOL = " & medMaxAmnt & ", "
        If Len(comSalDepn) > 0 Then SQLQ = SQLQ & "BF_SALARYDEPENDANT = '" & Left(comSalDepn, 1) & "', "
        If IsNumeric(medMaxCover) Then SQLQ = SQLQ & "BF_MAXIMUM = " & Val(medMaxCover) & ", "
        If IsNumeric(medMinCover) Then SQLQ = SQLQ & "BF_MINIMUM = " & Val(medMinCover) & ", "
        If IsNumeric(medSalFactor) Then SQLQ = SQLQ & "BF_FACTOR = " & Val(medSalFactor) & ", "
        If IsNumeric(comRndFactor) Then SQLQ = SQLQ & "BF_ROUND = " & Val(comRndFactor) & ", "
        SQLQ = SQLQ & "BF_NEXTNEAREST= '" & IIf(optRound(0), "R", "N") & "', "
    End If
    
    'Ticket #24220 - Change Cost Only
    If optBenefit(2) Then
        If IsNumeric(medPayPeriodAmount) Then SQLQ = SQLQ & "BF_PPAMT = " & medPayPeriodAmount & ", "
    Else
        SQLQ = SQLQ & "BF_PPAMT = " & medPayPeriodAmount & ", "
    End If
    
    SQLQ = SQLQ & "BF_PCE =  " & medPPE & ", "
    SQLQ = SQLQ & "BF_PCC =  " & medPPComp & ", "
    SQLQ = SQLQ & "BF_PREMIUM = '" & lblAP & "', "
    
    'Ticket #24220 - Change Cost Only
    If optBenefit(2) Then
        If IsNumeric(txtPer) Then SQLQ = SQLQ & "BF_PER =  " & txtPer & ",  "
        If IsNumeric(medUnitCost) Then SQLQ = SQLQ & "BF_UNITCOST = " & medUnitCost & ", "
    Else
        SQLQ = SQLQ & "BF_PER =  " & txtPer & ",  "
        SQLQ = SQLQ & "BF_UNITCOST = " & medUnitCost & ", "
    End If
    
    If txtTAXBEN <> "" Then SQLQ = SQLQ & "BF_TAXBEN = '" & txtTAXBEN & "',  "
    
    ' danielk - 02/05/2003 - Added pre/post tax
    SQLQ = SQLQ & "BF_PTAX = '" & txtPreAftTax & "', "
    ' danielk - 02/05/2003 - end
    
    If comSalDepn = "Yes" Then
        If glbSQL Or glbOracle Then
            SalaryStr = "(CASE WHEN HR_SALARY_HISTORY.SH_SALCD='A' THEN 1 WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN 12 " & _
                        "WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN " & strDays & " ELSE HR_SALARY_HISTORY.SH_WHRS*52 END) "
            If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                SalaryStr = SalaryStr & "* (CASE WHEN HR_SALARY_HISTORY.SH_NFAC_SALARY<>0 AND HR_SALARY_HISTORY.SH_NFAC_SALARY IS NOT NULL THEN HR_SALARY_HISTORY.SH_NFAC_SALARY ELSE HR_SALARY_HISTORY.SH_SALARY END) * " & Val(medSalFactor)
            Else
                SalaryStr = SalaryStr & "* HR_SALARY_HISTORY.SH_SALARY * " & Val(medSalFactor)
            End If
            SalaryStr = "(CASE WHEN " & Val(medMinCover) & "> 0 And " & Val(medMinCover) & " >" & SalaryStr & " Then " & Val(medMinCover) & " ELSE " & SalaryStr & " END)"
            SalaryStr = "(CASE WHEN " & Val(medMaxCover) & "> 0 And " & Val(medMaxCover) & " <" & SalaryStr & " Then " & Val(medMaxCover) & " ELSE " & SalaryStr & " END)"
            If Val(comRndFactor) > 0 Then SalaryStr = "ROUND(" & SalaryStr & "/" & Val(comRndFactor) & _
                IIf(optRound(0), "+0", "+.49") & ",0) *" & Val(comRndFactor)
        Else
            SalaryStr = "iif(HR_SALARY_HISTORY.SH_SALCD='M',HR_SALARY_HISTORY.SH_SALARY * 12, " & _
            "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52), " & _
            "iif(HR_SALARY_HISTORY.SH_SALCD='D',HR_SALARY_HISTORY.SH_SALARY * " & strDays & ", " & _
            "HR_SALARY_HISTORY.SH_SALARY))) * " & Val(medSalFactor)
            SalaryStr = "IIf(" & Val(medMinCover) & "> 0 And " & Val(medMinCover) & " >" & SalaryStr & "," & Val(medMinCover) & "," & SalaryStr & ")"
            SalaryStr = "IIf(" & Val(medMaxCover) & "> 0 And " & Val(medMaxCover) & " <" & SalaryStr & "," & Val(medMaxCover) & "," & SalaryStr & ")"
            If Val(comRndFactor) > 0 Then SalaryStr = "ROUND(" & SalaryStr & "/" & Val(comRndFactor) & _
                IIf(optRound(0), "+0", "+.5") & ") *" & Val(comRndFactor)
        End If
        SQLQ = SQLQ & "BF_AMT = " & SalaryStr & ",  "
    Else
        'Ticket #24220 - Change Cost Only
        If optBenefit(2) Then
            If IsNumeric(medCovAmount) Then SQLQ = SQLQ & "BF_AMT = " & medCovAmount & ",  "
        Else
            SQLQ = SQLQ & "BF_AMT = " & medCovAmount & ",  "
        End If
    End If
    
    If comSalDepn = "Yes" And optActual(1) Then
        TCostStr = "(" & SalaryStr & ")/" & txtPer & "*" & medUnitCost
            SQLQ = SQLQ & "BF_TCOST = " & TCostStr & ", "
        CostStr = TCostStr & "*" & medPPE
            SQLQ = SQLQ & "BF_ECOST = " & CostStr & ", "
        CostStr = CostStr & "/12"
            SQLQ = SQLQ & "BF_MTHECOST = " & CostStr & ", "
        CostStr = TCostStr & "*" & medPPComp
            SQLQ = SQLQ & "BF_CCOST = " & CostStr & ", "
        CostStr = CostStr & "/12"
            SQLQ = SQLQ & "BF_MTHCCOST = " & CostStr & ", "
    Else
        'Ticket #24220 - Change Cost Only
        If optBenefit(2) Then
            If IsNumeric(medEECost) Then SQLQ = SQLQ & "BF_ECOST = " & medEECost & ", "
            If IsNumeric(medCompCost) Then SQLQ = SQLQ & "BF_CCOST =  " & medCompCost & ", "
        
            If IsNumeric(medTCost) Then SQLQ = SQLQ & "BF_TCOST = " & medTCost & ", "
            
            If IsNumeric(medMCCOST) Then SQLQ = SQLQ & "BF_MTHCCOST = " & medMCCOST & ",  "
            If IsNumeric(medMECOST) Then SQLQ = SQLQ & "BF_MTHECOST = " & medMECOST & ",  "
            
        Else
            SQLQ = SQLQ & "BF_ECOST = " & medEECost & ", "
            SQLQ = SQLQ & "BF_CCOST =  " & medCompCost & ", "
        
            SQLQ = SQLQ & "BF_TCOST = " & medTCost & ", "
            
            SQLQ = SQLQ & "BF_MTHCCOST = " & medMCCOST & ",  "
            SQLQ = SQLQ & "BF_MTHECOST = " & medMECOST & ",  "
        End If
        
    End If
    SQLQ = SQLQ & "BF_LDATE = " & Date_SQL(Date) & ","
    SQLQ = SQLQ & "BF_LTIME = '" & Time$ & "', "
    SQLQ = SQLQ & "BF_LUSER = '999999998' "
    
    If glbSQL And comSalDepn = "Yes" Then
        SQLQ = SQLQ & "From HRBENFT LEFT JOIN HR_SALARY_HISTORY ON HRBENFT.BF_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR "
    End If
    SQLQ = SQLQ & " WHERE " & fglbWSQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
    If comSalDepn = "Yes" Then SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT<>0"
End If

XUpdCount = 0
If chkEmployee(0) Then
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    If Not AUDITBENF("M") Then MsgBox "ERROR - AUDIT FILE"
    
'    USQLQ = "SELECT BF_EMPNBR FROM HRBENFT WHERE BF_LUSER='999999998' "
'    RSEMPLIST.Open USQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
'    Do While Not RSEMPLIST.EOF
'        If Len(strEMPLIST) > 0 Then
'            strEMPLIST = strEMPLIST & "," & RSEMPLIST("BF_EMPNBR")
'        Else
'            strEMPLIST = strEMPLIST & RSEMPLIST("BF_EMPNBR")
'        End If
'        RSEMPLIST.MoveNext
'    Loop
'    RSEMPLIST.Close
    
    USQLQ = "UPDATE HRBENFT SET BF_LUSER='" & glbUserID & "' WHERE BF_LUSER='999999998' "
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute USQLQ, XUpdCount
    gdbAdoIhr001.CommitTrans
End If

If chkEmployee(1) Then
    SQLQ = Replace(SQLQ, "HREMP", "Term_HREMP")
    SQLQ = Replace(SQLQ, "HRBENFT", "Term_HRBENFT")
    SQLQ = Replace(SQLQ, "HR_SALARY_HISTORY", "Term_SALARY_HISTORY")
    'SQLQ = Replace(SQLQ, "999999998", glbUserID)
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ ', RecsAffected
    'XUpdCount = XUpdCount + RecsAffected
    gdbAdoIhr001.CommitTrans
    
    USQLQ = "UPDATE Term_HRBENFT SET BF_LUSER='" & glbUserID & "' WHERE BF_LUSER='999999998' "
    gdbAdoIhr001.BeginTrans
    'gdbAdoIhr001.Execute USQLQ, XUpdCount
    gdbAdoIhr001.Execute SQLQ, RecsAffected
    XUpdCount = XUpdCount + RecsAffected
    gdbAdoIhr001.CommitTrans
End If

Dim rsEMP As New ADODB.Recordset
Dim xEmpnbr

'Active Employees
If chkEmployee(0) Then
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & " "
    If Len(fglbWSQLQ) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT BF_EMPNBR FROM HRBENFT WHERE " & fglbWSQLQ & ")"
    End If

    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsEMP.EOF
        xEmpnbr = rsEMP("ED_EMPNBR")
        Call updBenefitForSalDEPN(xEmpnbr)
        If glbGP Then 'Ticket #30111 Franks 06/20/2017
            Call Employee_GP_NewBenefitDeduction_Integration(xEmpnbr)
        End If
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & rsEMP("ED_EMPNBR")
        Else
            strEMPLIST = rsEMP("ED_EMPNBR")
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End If

'Terminated Employees
If chkEmployee(1) Then
    SQLQ = "SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & " "
    If Len(fglbWSQLQ) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT BF_EMPNBR FROM Term_HRBENFT WHERE " & fglbWSQLQ & ")"
    End If
    
    rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsEMP.EOF
        xEmpnbr = rsEMP("ED_EMPNBR")
        'Call updBenefitForSalDEPN(xEMPNBR)
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & rsEMP("ED_EMPNBR")
        Else
            strEMPLIST = rsEMP("ED_EMPNBR")
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End If


modUpdRecs = True
Exit Function

modUpdRecs2_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdRecs", "Benefits Reason", "Update")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modUpdGroup() As Boolean
Dim SQLQ As String
Dim rsEMP As New ADODB.Recordset
Dim xEmpnbr
Dim OLDBGroup
Dim NewBGroup
Dim Msg$
Dim xTotalRecs

modUpdGroup = False
On Error GoTo modUpdGroup_Err

MDIMain.panHelp(0).FloodType = 1

xTotalRecs = 0
XUpdCount = 0

Screen.MousePointer = HOURGLASS

Call getWSQLQ("U")

fglbDupCode = False

If chkEmployee(0) Then
    SQLQ = "SELECT ED_EMPNBR,ED_BENEFIT_GROUP FROM HREMP WHERE " & fglbESQLQ
    If Len(clpBGroup(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_BENEFIT_GROUP='" & clpBGroup(0).Text & "'"
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    xTotalRecs = rsEMP.RecordCount
    Do Until rsEMP.EOF
        MDIMain.panHelp(0).FloodPercent = (XUpdCount / xTotalRecs) * 100
        XUpdCount = XUpdCount + 1
        xEmpnbr = rsEMP("ED_EMPNBR")
        OLDBGroup = rsEMP("ED_BENEFIT_GROUP")
        NewBGroup = Trim(clpBGroup(1).Text)
        If Not EmpHisCalc(0, xEmpnbr, "", "", "", "", "", "", "", Date, NewBGroup) Then MsgBox "EMPHIS Error"
        rsEMP("ED_BENEFIT_GROUP") = NewBGroup
        rsEMP.Update
        Call UPDBenefit(xEmpnbr, NewBGroup, "A")
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & rsEMP("ED_EMPNBR")
        Else
            strEMPLIST = rsEMP("ED_EMPNBR")
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End If


If chkEmployee(1) Then
    SQLQ = "SELECT ED_EMPNBR,ED_BENEFIT_GROUP FROM Term_HREMP WHERE " & fglbESQLQ
    If Len(clpBGroup(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_BENEFIT_GROUP='" & clpBGroup(0).Text & "'"
    rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    xTotalRecs = xTotalRecs + rsEMP.RecordCount
    Do Until rsEMP.EOF
        MDIMain.panHelp(0).FloodPercent = (XUpdCount / xTotalRecs) * 100
        XUpdCount = XUpdCount + 1
        xEmpnbr = rsEMP("ED_EMPNBR")
        If IsNull(rsEMP("ED_BENEFIT_GROUP")) Then OLDBGroup = "" Else OLDBGroup = rsEMP("ED_BENEFIT_GROUP")
        NewBGroup = Trim(clpBGroup(1).Text)
        rsEMP("ED_BENEFIT_GROUP") = NewBGroup
        rsEMP.Update
        Call UPDBenefit(xEmpnbr, NewBGroup, "T")
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & rsEMP("ED_EMPNBR")
        Else
            strEMPLIST = rsEMP("ED_EMPNBR")
        End If
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End If

MDIMain.panHelp(0).FloodType = 0
If fglbDupCode Then
    Msg$ = "The Benefit Group contains Benefits with multiple coverages. " & Chr(10)
    Msg$ = Msg$ & "Edit individual employee Benefit records to ensure that " & Chr(10)
    Msg$ = Msg$ & "the employee has the appropriate coverage."
    MsgBox Msg$
End If
modUpdGroup = True
Exit Function

modUpdGroup_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdGroup", "Benefits Group", "Update")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub UPDBenefit(xEmpnbr, NewBGroup, TermOrActive)
Dim rsBGMST As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim xCode
Dim SQLQ As String
Dim xACT

SQLQ = "SELECT BM_BCODE FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "' ORDER BY BM_BCODE "
rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
xCode = ""

Do While Not rsBGMST.EOF
    If xCode = rsBGMST("BM_BCODE") Then
        fglbDupCode = True
    End If
    xCode = rsBGMST("BM_BCODE")
    rsBGMST.MoveNext
Loop
rsBGMST.Close

Call updateBenefit(xEmpnbr, NewBGroup, TermOrActive, MassUpdateBenefitGroup)

If glbGP Then 'Ticket #30111 Franks 06/13/2017
    Call Employee_GP_NewBenefitDeduction_Integration(xEmpnbr)
End If

'Call updBenefitForSalDEPN(XEMPNBR)
End Sub

Private Function getWSQLQ(FSTR) As String

fglbESQLQ = glbSeleDeptUn

If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_REGION = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "' "
If Len(clpCode(6).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ADMINBY = '" & clpCode(6).Text & "' "
If Len(clpCode(7).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_BENEFIT_GROUP = '" & clpCode(7).Text & "' "
If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

If FSTR <> "A" Then
    If glbLinamar Then
        fglbWSQLQ = " BF_BCODE = '" & IIf(clpDiv.Text = "", "ALL", clpDiv.Text) & clpCode(4).Text & "' "
    Else
        fglbWSQLQ = " BF_BCODE = '" & clpCode(4).Text & "' "
    End If
    If Len(txtCovType) > 0 Then
        fglbWSQLQ = fglbWSQLQ & " AND BF_COVER= '" & txtCovType & "'"
    End If
    If FSTR = "D" And dlpEDate.Text <> "" Then
        fglbWSQLQ = fglbWSQLQ & " AND BF_EDATE=" & Date_SQL(dlpEDate.Text)
    End If
End If

End Function

Private Function getWSQLQRPT() As String

'getWSQLQRPT = glbSeleDeptUn    'Department security removed by Bryan, redundant, this is a list of changes, whether they have security is irrelevant at this point
'If Len(clpDept.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DEPTNO} = '" & clpDept.Text & "')"
'If Len(clpDiv.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DIV} = '" & clpDiv.Text & "') "
'If Len(clpCode(1).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_LOC} = '" & clpCode(1).Text & "') "
'If Len(clpCode(2).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ORG} = '" & clpCode(2).Text & "') "
'If Len(clpCode(3).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_EMP} = '" & clpCode(3).Text & "') "
'If Len(clpCode(5).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_REGION} = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "') "
'If Len(clpCode(6).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ADMINBY} = '" & clpCode(6).Text & "') "
'If Len(clpCode(7).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_BENEFIT_GROUP} = '" & clpCode(7).Text & "') "
'If Len(clpPT.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_PT} = '" & clpPT.Text & "') "
If Len(strEMPLIST) > 0 Then getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) "

End Function

Private Sub medUnitCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optActual_Click(Index As Integer, Value As Integer)
If optActual(1).Value = True Then
    txtPer.Enabled = True
    medUnitCost.Enabled = True
    'medCovAmount.Enabled = True  'jaddy 11/2/99
    medTCost = "From System"
    medTCost.Enabled = False
Else
    'txtPer = 0
    txtPer.Enabled = False
    'medUnitCost = 0
    medUnitCost.Enabled = False
    'medCovAmount.Enabled = False  'jaddy 11/2/99
    medTCost.Enabled = True
    medTCost = ""
End If
End Sub

Private Sub optActual_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optActual_LostFocus(Index As Integer)
If optActual(0).Value = True Then
    lblAP = "A"
Else
    lblAP = "P"
End If

End Sub

Private Sub optBenefit_Click(Index As Integer)
frmMaster.Visible = False
frmGroup.Visible = False
If optBenefit(0) Or optBenefit(2) Then  'Ticket #24220 - Change Cost Only
    frmMaster.Visible = True
Else
    frmGroup.Visible = True
End If
Call SET_UP_MODE
End Sub

Private Sub txtCovType_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtPer_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPer_Change()
Call medCovAmount_Change
End Sub

Private Sub txtTAXBEN_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub CalcBeneCasey()
    Dim Total As Double, ECost As Double, CCost As Double
    
    If ChangingFields Then Exit Sub
    ChangingFields = True   ' Flag to tell other code we're changing fields, prevents loop
    ECost = Val(medMECOST.Text)
    CCost = Val(medMCCOST.Text)
    Total = ECost + CCost
    If Total = 0 Then
        medPPE.Text = "0"
        medPPComp.Text = "0"
        medEECost.Text = "0"
        medCompCost.Text = "0"
    Else
        medPPE.Text = ECost / Total
        medPPComp.Text = CCost / Total
        medEECost.Text = ECost * 12
        medCompCost.Text = CCost * 12
        medTCost.Text = Total * 12
    End If
    ChangingFields = False
End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
'UpdateRight = gSec_Upd_Benefits
UpdateRight = GetMassUpdateSecurities("Benefits_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = optBenefit(0)
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = optBenefit(0)
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Sub txtTAXBEN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim rsBenfit As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0
    
    Call getWSQLQ("D")

    If chkEmployee(0) Then
        SQLQ = "SELECT COUNT(BF_EMPNBR) AS TOT_REC FROM HRBENFT "
        SQLQ = SQLQ & " WHERE " & fglbWSQLQ
        SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & " )"
        rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsBenfit.EOF Then
            recCount = rsBenfit("TOT_REC")
        Else
            recCount = 0
        End If
        rsBenfit.Close
        Set rsBenfit = Nothing
    End If
        
    If chkEmployee(1) Then
        SQLQ = "SELECT COUNT(BF_EMPNBR) AS TOT_REC FROM Term_HRBENFT"
        SQLQ = SQLQ & " WHERE " & fglbWSQLQ
        SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & " )"
        rsBenfit.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not rsBenfit.EOF Then
            recCount = recCount + rsBenfit("TOT_REC")
        Else
            recCount = recCount + 0
        End If
        rsBenfit.Close
        Set rsBenfit = Nothing
    End If
    
    getRecordCount_Delete = recCount
    
End Function

Private Function getRecordCount_ModifyBenGrp()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_ModifyBenGrp = 0
    recCount = 0
    
    Call getWSQLQ("U")

    If chkEmployee(0) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP WHERE " & fglbESQLQ
        If Len(clpBGroup(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_BENEFIT_GROUP='" & clpBGroup(0).Text & "'"
        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        If Not rsEMP.EOF Then
            recCount = rsEMP("TOT_REC")
        Else
            recCount = 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
        
    If chkEmployee(1) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM Term_HREMP WHERE " & fglbESQLQ
        If Len(clpBGroup(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_BENEFIT_GROUP='" & clpBGroup(0).Text & "'"
        rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
        If Not rsEMP.EOF Then
            recCount = recCount + rsEMP("TOT_REC")
        Else
            recCount = recCount + 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
    
    getRecordCount_ModifyBenGrp = recCount
    
End Function

Private Function getRecordCount_ModifyBenMst()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_ModifyBenMst = 0
    recCount = 0
    
    Call getWSQLQ("U")

    If chkEmployee(0) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP WHERE " & fglbESQLQ & " "
        If Len(fglbWSQLQ) > 0 Then
            SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT BF_EMPNBR FROM HRBENFT WHERE " & fglbWSQLQ & ")"
        End If
        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEMP.EOF Then
            recCount = rsEMP("TOT_REC")
        Else
            recCount = 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
        
    If chkEmployee(1) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM Term_HREMP WHERE " & fglbESQLQ & " "
        If Len(fglbWSQLQ) > 0 Then
            SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT BF_EMPNBR FROM Term_HRBENFT WHERE " & fglbWSQLQ & ")"
        End If
        rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not rsEMP.EOF Then
            recCount = recCount + rsEMP("TOT_REC")
        Else
            recCount = recCount + 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
    
    getRecordCount_ModifyBenMst = recCount
    
End Function

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0
    
    Call getWSQLQ("A")

    If chkEmployee(0) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP WHERE " & fglbESQLQ & " "
        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEMP.EOF Then
            recCount = rsEMP("TOT_REC")
        Else
            recCount = 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
        
    If chkEmployee(1) Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM TERM_HREMP WHERE " & fglbESQLQ & " "
        rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not rsEMP.EOF Then
            recCount = recCount + rsEMP("TOT_REC")
        Else
            recCount = recCount + 0
        End If
        rsEMP.Close
        Set rsEMP = Nothing
    End If
    
    getRecordCount_Add = recCount

End Function

Private Sub EmailNotification(xUpdType)
    Dim rsEmpBen As New ADODB.Recordset
    Dim SQLQ As String
    
    If xUpdType = "ADD" Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR IN (" & EmpListAdd & ")"
        If glbLinamar Then
            SQLQ = SQLQ & " AND BF_BCODE = '" & IIf(clpDiv.Text = "", "ALL", clpDiv.Text) & clpCode(4).Text & "' "
        Else
            SQLQ = SQLQ & " AND BF_BCODE = '" & clpCode(4).Text & "' "
        End If
        'If Len(txtCovType) > 0 Then
            SQLQ = SQLQ & " AND  BF_COVER= '" & txtCovType & "'"
        'End If
    Else
        SQLQ = "SELECT * FROM HRBENFT WHERE 1=1"
        If xUpdType <> "ADD" Then
            SQLQ = SQLQ & " AND " & fglbWSQLQ & " "
        End If
        If Len(fglbESQLQ) > 0 Then
            SQLQ = SQLQ & " AND BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        End If
    End If
    
    rsEmpBen.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsEmpBen.EOF
        'Benefits Added
        If xUpdType = "ADD" Then
            MailBody = "The New Benefit:" & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & rsEmpBen("BF_EMPNBR") & vbCrLf
            MailBody = MailBody & "Name: " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_SURNAME") & ", " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_FNAME") & vbCrLf
            MailBody = MailBody & "New Benefit: " & GetTABLDesc("BNCD", rsEmpBen("BF_BCODE")) & vbCrLf
            MailBody = MailBody & "Effective Date: " & Format(CVDate(rsEmpBen("BF_EDATE")), "SHORT DATE") & vbCrLf
        End If
                                
        'Benefits Updated
        If xUpdType = "UPDATE" Then
            MailBody = "The Updated Benefit:" & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & rsEmpBen("BF_EMPNBR") & vbCrLf
            MailBody = MailBody & "Name: " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_SURNAME") & ", " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_FNAME") & vbCrLf
            MailBody = MailBody & "Updated Benefit: " & GetTABLDesc("BNCD", rsEmpBen("BF_BCODE")) & vbCrLf
            MailBody = MailBody & "Effective Date: " & Format(CVDate(rsEmpBen("BF_EDATE")), "SHORT DATE") & vbCrLf
        End If
        
        'Benefits Deleted
        If xUpdType = "DELETE" Then
            MailBody = "The Deleted Benefit:" & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & rsEmpBen("BF_EMPNBR") & vbCrLf
            MailBody = MailBody & "Name: " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_SURNAME") & ", " & GetEmpData(rsEmpBen("BF_EMPNBR"), "ED_FNAME") & vbCrLf
            MailBody = MailBody & "Deleted Benefit: " & GetTABLDesc("BNCD", rsEmpBen("BF_BCODE")) & vbCrLf
            MailBody = MailBody & "Effective Date: " & Format(CVDate(rsEmpBen("BF_EDATE")), "SHORT DATE") & vbCrLf
        End If
        
        Call imgEmail_ClickX(xUpdType, rsEmpBen("BF_EMPNBR"))
        
        rsEmpBen.MoveNext
    Loop
    rsEmpBen.Close
    Set rsEmpBen = Nothing
    
    Screen.MousePointer = DEFAULT

End Sub

Public Sub imgEmail_ClickX(xType, xEmpnbr)
Dim xEmail
Dim xToEmail As String

On Error GoTo Email_Err

        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        
        If Len(xEmail) > 0 Then
            'Ticket #18090 - begin
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            Else
                'Ticket #20317 - More Emails for everyone
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            End If
            'Ticket #18090 - end
            
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONBENEFIT")
            
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
            Else
                frmSendEmail.txtCC.Text = xEmail
            End If
                        
            'Email Subject line based on the Type of Email
            If xType = "DELETE" Then
                frmSendEmail.txtSubject.Text = "info:HR Benefit Delete Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            ElseIf xType = "UPDATE" Then
                frmSendEmail.txtSubject.Text = "info:HR Benefit Update Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            Else
                frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            End If
            frmSendEmail.txtBody.Text = MailBody
            
            'Not showing the Email Send window as this is a mass update
            'frmSendEmail.Show 1
            frmSendEmail.cmdSend_Click
        Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & GetEmpData(xEmpNbr, "ED_SURNAME") & ", " & GetEmpData(xEmpNbr, "ED_FNAME")
            'End If
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

