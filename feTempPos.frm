VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmETmpCrsTrnPos 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Temporary or Cross Training Assignment"
   ClientHeight    =   9330
   ClientLeft      =   240
   ClientTop       =   735
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9330
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feTempPos.frx":0000
      Height          =   1725
      Left            =   0
      OleObjectBlob   =   "feTempPos.frx":0014
      TabIndex        =   0
      Top             =   510
      Width           =   11415
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5295
      LargeChange     =   315
      Left            =   11310
      Max             =   100
      SmallChange     =   315
      TabIndex        =   36
      Top             =   2400
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   45
      Top             =   8790
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   952
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         TabIndex        =   94
         Tag             =   "Job Files related to this Job"
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
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
         TabIndex        =   39
         Tag             =   "Call Performance Form"
         Top             =   330
         Visible         =   0   'False
         Width           =   1250
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
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
         TabIndex        =   93
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   6030
         Width           =   1005
      End
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   0
      TabIndex        =   46
      Top             =   2280
      Width           =   11355
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   7020
         TabIndex        =   95
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame frmJobEnd 
         Height          =   615
         Left            =   5100
         TabIndex        =   90
         Top             =   3070
         Width           =   5415
         Begin INFOHR_Controls.DateLookup dlpENDDATE 
            DataField       =   "TW_ENDDATE"
            Height          =   285
            Left            =   1170
            TabIndex        =   29
            Tag             =   "41-Temp./Cross Training Position End Date"
            Top             =   5
            Width           =   2800
            _ExtentX        =   4948
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "TW_ENDREAS"
            Height          =   285
            Index           =   2
            Left            =   1170
            TabIndex        =   30
            Tag             =   "01-End Reason Code"
            Top             =   300
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
            TabIndex        =   92
            Top             =   305
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
            TabIndex        =   91
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.ComboBox cboShift 
         Height          =   315
         Left            =   2880
         TabIndex        =   89
         Top             =   3090
         Visible         =   0   'False
         Width           =   855
      End
      Begin Threed.SSCheck chkActPosition 
         Height          =   255
         Left            =   5340
         TabIndex        =   88
         Top             =   30
         Visible         =   0   'False
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
      Begin VB.TextBox txtComments2 
         Appearance      =   0  'Flat
         DataField       =   "TW_COMMENT2"
         Height          =   285
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "00-Temp./Cross Training Position Notes 2"
         Top             =   4740
         Width           =   2895
      End
      Begin VB.Frame frmMulti 
         Height          =   3195
         Left            =   4980
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CheckBox chkUseForBenefit 
            Caption         =   "For Benefit"
            Height          =   315
            Left            =   3360
            TabIndex        =   87
            Top             =   2820
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtPayrollID 
            Appearance      =   0  'Flat
            DataField       =   "TW_PAYROLL_ID"
            Height          =   285
            Left            =   1600
            MaxLength       =   25
            TabIndex        =   31
            Tag             =   "00-Payroll ID"
            Top             =   2850
            Width           =   1680
         End
         Begin VB.TextBox txtEmpType 
            BackColor       =   &H80000004&
            DataField       =   "TW_LEADHAND"
            Height          =   285
            Left            =   3960
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1950
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox comEmpType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1610
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Tag             =   "10-Type of Employee "
            Top             =   1920
            Visible         =   0   'False
            Width           =   2800
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "TW_EMP"
            Height          =   285
            Index           =   4
            Left            =   1290
            TabIndex        =   23
            Tag             =   "00-Employment Status - Code"
            Top             =   1020
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDEM"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "TW_ORG"
            Height          =   285
            Index           =   0
            Left            =   1290
            TabIndex        =   24
            Tag             =   "00-Union - Code"
            Top             =   1320
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin INFOHR_Controls.CodeLookup clpDiv 
            DataField       =   "TW_DIV"
            Height          =   285
            Left            =   1290
            TabIndex        =   20
            Tag             =   "00-Specific Division Desired"
            Top             =   120
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   1
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            DataField       =   "TW_DEPTNO"
            Height          =   285
            Left            =   1290
            TabIndex        =   21
            Top             =   420
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   7
            LookupType      =   2
         End
         Begin INFOHR_Controls.CodeLookup clpGLNum 
            DataField       =   "TW_GLNO"
            Height          =   285
            Left            =   1290
            TabIndex        =   22
            Tag             =   "00-General Ledger - Code"
            Top             =   720
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "TW_SECTION"
            Height          =   285
            Index           =   5
            Left            =   1290
            TabIndex        =   27
            Tag             =   "00-Section - Code"
            Top             =   1920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpPT 
            DataField       =   "TW_PT"
            DataSource      =   " "
            Height          =   285
            Left            =   1290
            TabIndex        =   25
            Tag             =   "00-Category Codes"
            Top             =   1620
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDPT"
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
            TabIndex        =   81
            Top             =   2880
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
            TabIndex        =   80
            Top             =   1620
            Width           =   765
         End
         Begin VB.Label lblSection 
            AutoSize        =   -1  'True
            Caption         =   "Section"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1945
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
            TabIndex        =   75
            Top             =   1020
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
            TabIndex        =   74
            Top             =   720
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
            TabIndex        =   73
            Top             =   420
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
            TabIndex        =   72
            Top             =   150
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
            TabIndex        =   51
            Top             =   1320
            Width           =   510
         End
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         DataField       =   "TW_GRID"
         Height          =   285
         Left            =   1590
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGD"
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         DataField       =   "TW_SDATE"
         Height          =   285
         Left            =   1590
         TabIndex        =   4
         Tag             =   "41-Temp./Cross Training Position Start Date"
         Top             =   780
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "TW_SHIFT"
         Height          =   285
         Left            =   1905
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "00-Code assigned to the shift"
         Top             =   3090
         Width           =   810
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "TW_LDATE"
         Height          =   285
         Index           =   0
         Left            =   10320
         MaxLength       =   25
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Ldate"
         Top             =   4470
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "TW_LTIME"
         Height          =   285
         Index           =   1
         Left            =   10920
         MaxLength       =   25
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "LTime"
         Top             =   4470
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "TW_LUSER"
         Height          =   285
         Index           =   2
         Left            =   11520
         MaxLength       =   25
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "LUser"
         Top             =   4110
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         DataField       =   "TW_COMMENT"
         Height          =   285
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "00-Temp./Cross Training Position Notes 1"
         Top             =   4410
         Width           =   2895
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "TW_REPTAU"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   35
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1110
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "TW_REPTAU2"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   37
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "TW_REPTAU3"
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   38
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1770
         Visible         =   0   'False
         Width           =   1215
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
         DataField       =   "TW_DHRS"
         Height          =   288
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Tag             =   "10-Usual working hours per day"
         Top             =   2100
         Width           =   876
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
         DataField       =   "TW_WHRS"
         Height          =   285
         Index           =   1
         Left            =   1905
         TabIndex        =   9
         Tag             =   "10- Number of hours in work week"
         Top             =   2430
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
         DataField       =   "TW_PHRS"
         Height          =   285
         Index           =   2
         Left            =   1905
         TabIndex        =   10
         Tag             =   "10-Usual working hours per pay period"
         Top             =   2760
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
         DataField       =   "TW_FTENUM"
         Height          =   285
         Left            =   1905
         TabIndex        =   13
         Tag             =   "10-Full - time equivalency"
         Top             =   3750
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
         Left            =   2400
         TabIndex        =   55
         Top             =   3750
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
         DataField       =   "TW_CURRENT"
         Height          =   285
         Index           =   0
         Left            =   7200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Current Temporary/Cross Training Position Record"
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
         DataField       =   "TW_FTEHRS"
         Height          =   285
         Left            =   1905
         TabIndex        =   14
         Tag             =   "10-FTE Hours worked per year"
         Top             =   4080
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
         DataField       =   "TW_JREASON"
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   12
         Tag             =   "01-Reason for change in Temp./Cross Training position - Code"
         Top             =   3420
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   7
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
         Left            =   1590
         TabIndex        =   6
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
         Left            =   1590
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
         DataField       =   "TW_JOB"
         Height          =   285
         Left            =   1590
         TabIndex        =   1
         Tag             =   "01-Temporary/Cross Training Position code"
         Top             =   120
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   7320
         TabIndex        =   18
         Tag             =   "00-Band - Code"
         Top             =   3450
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin INFOHR_Controls.CodeLookup clpPayrollCategory 
         DataField       =   "TW_PAYROLL_CATEGORY"
         Height          =   285
         Left            =   6720
         TabIndex        =   32
         Top             =   4440
         Visible         =   0   'False
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   9
      End
      Begin VB.Frame frmOCCAC 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -60
         TabIndex        =   78
         Top             =   5070
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox txtPosCtr 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TW_POSITION_CONTROL"
            Height          =   285
            Left            =   1970
            MaxLength       =   6
            TabIndex        =   17
            Tag             =   "00-CCAC Temporary/Cross Training Position #"
            Top             =   30
            Width           =   1155
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   1710
            Picture         =   "feTempPos.frx":7594
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lblPosCtr 
            Caption         =   "CCAC Position #"
            Height          =   345
            Left            =   120
            TabIndex        =   79
            Top             =   30
            Width           =   1575
         End
      End
      Begin VB.Frame frmLinamar 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   47
         Top             =   5100
         Visible         =   0   'False
         Width           =   3405
         Begin VB.CheckBox chkLeadHand 
            Height          =   195
            Left            =   1755
            TabIndex        =   19
            Tag             =   "40-Lead Hand - y/n"
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblLeadHand 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "TW_LeadHand"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2640
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   0
            Width           =   795
         End
      End
      Begin MSMask.MaskEdBox medBillingRate 
         Height          =   285
         Left            =   7020
         TabIndex        =   34
         Tag             =   "10-Enter Billing Rate"
         Top             =   5100
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
      Begin Threed.SSCheck chkTrackCrsRenewal 
         DataField       =   "TW_TRK_CRS_RENEWAL"
         Height          =   285
         Left            =   7200
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
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
         Left            =   5250
         TabIndex        =   96
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image imgNoSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   6360
         Picture         =   "feTempPos.frx":76DE
         Top             =   4080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   6360
         Picture         =   "feTempPos.frx":7828
         Top             =   4080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblBillingRate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Rate"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5250
         TabIndex        =   86
         Top             =   5100
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblComment2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 2"
         Height          =   255
         Left            =   60
         TabIndex        =   85
         Top             =   4740
         Width           =   855
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Occupation"
         Height          =   195
         Left            =   5250
         TabIndex        =   84
         Top             =   4800
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
         Left            =   7020
         TabIndex        =   33
         Top             =   4770
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
         Left            =   5250
         TabIndex        =   83
         Top             =   4500
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
         TabIndex        =   82
         Top             =   480
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for"
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
         TabIndex        =   77
         Top             =   3300
         Width           =   1050
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
         TabIndex        =   71
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
         TabIndex        =   70
         Top             =   810
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
         TabIndex        =   69
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   68
         Top             =   2130
         Width           =   930
      End
      Begin VB.Label lblHrsWeek 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Week"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   67
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lblHrsPayPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   66
         Top             =   2790
         Width           =   1515
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   65
         Top             =   3090
         Width           =   1725
      End
      Begin VB.Label lblEEStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
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
         TabIndex        =   64
         Top             =   3480
         Width           =   1050
      End
      Begin VB.Label lblFTENum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE#"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   63
         Top             =   3780
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
         TabIndex        =   62
         Top             =   4110
         Width           =   1395
      End
      Begin VB.Label lblCompNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CompNo"
         DataField       =   "TW_COMPNO"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8400
         TabIndex        =   61
         Top             =   3180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblComment 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 1"
         Height          =   255
         Left            =   60
         TabIndex        =   60
         Top             =   4410
         Width           =   855
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   1470
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
         TabIndex        =   57
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataField       =   "TW_EMPNBR"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9180
         TabIndex        =   56
         Top             =   3180
         Visible         =   0   'False
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmETmpCrsTrnPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim JobSnap_PayScale(15) As Double
Dim JobSnap_Salary_Code$
Dim JobSnap_MidPoint!
Dim fglbNew%
Dim savWHRS, savGrid, SavFte, SavFteHr, SavRpta(3), savSDate, savJOB
Dim Action

Dim fgtxtjob As String, fgtxtStartDate  As Variant
Dim fgtxtDhrs

Dim oPHRS, oWHRS, ODHRS, oJob, OSDATE
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
Dim oTrkCrsRen, flgNewCancel, savCurrent, flgTrainLstReset As Boolean

'frmMulti screen position: Left 4980 top 840

Private Function AUDITPSTN(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xPT, xDiv
Dim HRChangs As New Collection
Dim UpdateAudit As Boolean
Dim UptPositionDate As Date

On Error GoTo AUDIT_ERR

AUDITPSTN = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
Else
    xPT = ""
    xDiv = ""
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
        If isChanged_Field(HRChangs, oPayCategory, clpPayrollCategory, False) Then UpdateAudit = True
        Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
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
    If isChanged_Field(HRChangs, OLabourCD, clpCode(3)) Then UpdateAudit = True
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
            If chkCurrent(0) Or fglbNew Then   'Ticket #15751 - To prevent from updating Vadim with non-current position information.
                Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
            End If
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
rsTA("AU_DIVUPL") = xDiv
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
If OLabourCD <> clpCode(3).Text Then rsTA("AU_LABOURCD") = clpCode(3).Text

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID


rsTA("AU_LDATE") = Date
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
    SQLQ = "Select TW_JOB,TW_EMPNBR,TW_ID FROM HR_TEMP_WORK"
    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & " AND TW_CURRENT <>0 "
    SQLQ = SQLQ & " AND TW_PAYROLL_ID ='" & txtPayrollID & "'"
    SQLQ = SQLQ & " AND TW_ID <> " & Data1.Recordset!TW_ID
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job History Snap", "HR_TEMP_WORK", "SELECT")
Call RollBack '26July99 js
End Function

Private Function ifExistVadimPayrollID()
Dim x
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
Dim dd As Integer, DgDef As Double, Msg$, DCurSDate, DPrvSDate As Variant
Dim Response%, x%
Dim rsEmp As New ADODB.Recordset
Dim CaseyFlag As Boolean
chkPosition = False

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
    If Not glbCompSerial = "S/N - 2242W" And Not glbCompSerial = "S/N - 2390W" Then 'london ccac 'Collectcorp
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
'Ticket #16189-------------------------------
'If glbVadim Then
'    If glbMulti Then 'Ticket# 7751
'        If Len(txtPayrollID.Text) = 0 Then
'            MsgBox "Payroll ID is required"
'            txtPayrollID.SetFocus
'            Exit Function
'        End If
'    End If
'    If Len(clpPayrollCategory.Text) < 1 Then
'        MsgBox "Payroll Category is required field"
'        clpPayrollCategory.SetFocus
'        Exit Function
'    End If
'End If
'Ticket #16189-------------------------------

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
                    Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current Temporary/Cross Training position, "
                    Msg$ = Msg$ & Chr(10) & "go to Position screen under Employee Menu."
                    MsgBox Msg$
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
                        Msg$ = "Warning...you are creating a new position  for this employee."
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
                        MsgBox ("Shift is a required field")
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
    If Len(txtShift) < 1 Then
        MsgBox "Shift is a required field"
        txtShift.SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) < 1 Then
        MsgBox "Labour Code is required field"
        clpCode(3).SetFocus
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
    If Not glbLambton Then 'Ticket# 6692
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
For x% = 0 To 2
    If elpReptAuthShow(x%) = "0" Then elpReptAuthShow(x%) = ""
    If Len(elpReptAuthShow(x%)) > 0 Then
        If elpReptAuthShow(x%).Caption = "Unassigned" Then
            MsgBox "Rept. Authority Employee # not valid. Check Employee # and re-enter!"
            elpReptAuthShow(x%).SetFocus
            Exit Function
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
            If glbCompSerial <> "S/N - 2217W" Then ' Except CITY OF PICKERING
                chkCurrent(0) = False
            End If
            If DateDiff("d", dlpENDDATE.Text, dlpStartDate.Text) > 0 Then
                MsgBox "End Date Must Be Later than Start Date"
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
    'Ticket #16189-------------------------------
'    If glbVadim And glbMulti Then
'        If Not ChkTermPos Then Exit Function
'    End If
    'Ticket #16189-------------------------------
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
    If clpCode(0).Caption = "Unassigned" And Len(clpCode(0).Text) > 0 Then
        MsgBox lStr("Union Code must be valid")
        clpCode(0).SetFocus
        Exit Function
    End If
    If clpPT.Caption = "Unassigned" And Len(clpPT.Text) > 0 Then
        MsgBox lStr("Category Code must be valid")
        clpPT.SetFocus
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

For x% = 0 To 2
    If Not IsNumeric(medHours(x%)) Then medHours(x%) = 0
Next

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    If (savSDate <> dlpStartDate.Text And Not fglbNew) Or fglbNew Then
        If Not fglbNew And Not chkCurrent(0) Then
            DCurSDate = CurSDate()
            If DCurSDate > 0 Then    '0 if no current record out there
                DCurSDate = CVDate(DCurSDate)
                If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) <= 0 Then
                    MsgBox "Start Date cannot be same or later than your most current position record"
                    dlpStartDate.SetFocus
                    Exit Function
                End If
            End If
        End If
    
        DPrvSDate = PrvSDate(IIf(chkCurrent(0) And fglbNew, False, chkCurrent(0).Value))
        If DPrvSDate > 0 Then    '0 if no current record out there
            If DateDiff("d", DPrvSDate, CVDate(dlpStartDate.Text)) <= 0 Then
                MsgBox "Start Date cannot be same as or earlier than previous position(s)"
                dlpStartDate.SetFocus
                Exit Function
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
    
        'When trying to save an existing record which is not current
        If chkCurrent(0).Value = False And Not fglbNew Then
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
            'There should not be End Date for Current Position record
            If Len(Trim(dlpENDDATE.Text)) > 0 Then
                MsgBox "There should not be End Date for Current marked Position"
                dlpENDDATE.SetFocus
                Exit Function
            ElseIf Len(Trim(clpCode(2).Text)) > 0 Then
                MsgBox "There should not be End Reason for Current marked Position"
                clpCode(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If

'Ticket #16189-------------------------------
'If glbVadim Then
'    If Not AUDITPSTN(Action) Then MsgBox "ERROR - AUDIT FILE"
'Else
'    If DCurSDate = 0 Then DCurSDate = dlpStartDate.Text  'New Record
'    If IsDate(DCurSDate) Then  'Update Audit if Current Salary
'        If DateDiff("d", CVDate(dlpStartDate.Text), DCurSDate) <= 0 Then
'            If Not AUDITPSTN(Action) Then MsgBox "ERROR - AUDIT FILE"
'        End If
'    End If
'End If
'Ticket #16189-------------------------------


If glbWFC Then
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
End If

chkPosition = True

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
End Sub

Private Sub clpJob_LostFocus()
Call Job_Desc

'Ticket #16212 - Remove this logic because on Position Master it contains Hour/Pay Period
'If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
'    'Get the Hours per Day from HRJOB
'    medHours(0).Text = Get_DayHours_for_Job(clpJob.Text)
'End If

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

Private Sub cmdJobFiles_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpENDDATE_LostFocus()
    If Len(Trim(dlpENDDATE)) > 0 Then     'New Tracking method
        If IsDate(dlpENDDATE) Then
            chkTrackCrsRenewal.Visible = True
            chkCurrent(0).Value = False
            chkTrackCrsRenewal.Value = True
            'oTrkCrsRen = chkTrackCrsRenewal.Value
        Else
            MsgBox "Invalid End Date"
            dlpENDDATE.SetFocus
            Exit Sub
        End If
    Else
        If chkCurrent(0) = True Then
            chkTrackCrsRenewal.Value = False
        End If
    End If
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmETmpCrsTrnPos")
    Call FillMemoFile(SQLQ, "Offer")
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Offer"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmETmpCrsTrnPos")
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
    fraDetail.Height = 5535 '5000
    If glbLinamar Then fraDetail.Height = 7815
    If glbOttawaCCAC Then fraDetail.Height = 6000
    If Me.Height >= vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + panControls.Height Then '+ 230 Then
        scrControl.Value = 0
        fraDetail.Top = vbxTrueGrid.Height + panEEDESC.Height + 240
        scrControl.Visible = False
        Exit Sub
    End If
    If Me.Height < vbxTrueGrid.Height + panEEDESC.Height + scrControl.Top + panControls.Height + 400 Then Exit Sub
    scrControl.Visible = True
    
    scrControl.Max = vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + panControls.Height + 250 - Me.Height
    scrControl.Left = Me.Width - scrControl.Width - 120
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
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
    Dim x As Integer
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
    
    For x = 0 To 2
        Call txtReptAuthority_Change(x)
    Next
    
    'Call ST_UPD_MODE(True)  ' reset screen's attributes
    'Call SET_UP_MODE
    
    DoEvents
    
    If flgNewCancel And chkCurrent(0) Then
        'Changed 2
        'Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        Call Update_Employee_Job_Training_List(clpJob.Text, "Temporary")
        
        'Update records with tracking on
        rsDATA("TW_TRK_CRS_RENEWAL") = False
        rsDATA.Update
        
        chkTrackCrsRenewal.Value = False
    End If
    flgNewCancel = False
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdJobFiles.Enabled = False
        Else
            cmdJobFiles.Enabled = True
            
            If Not gSec_Inq_Job_Files_Attachment Then
                cmdJobFiles.Enabled = False
            End If
        End If
    End If
    
Exit Sub
Can_Err:
    If Err = 3018 Then
        Err = 0
        Resume Next
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_TEMP_WORK", "Cancel")
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
    'Ticket #16189-------------------------------
    'Call NextForm
    'Ticket #16189-------------------------------
    
    Unload Me
    If glbOnTop = "FRMETMPCRSTRNPOS" Then glbOnTop = ""
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
Dim xCurJob As String

    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
       MsgBox "Nothing to Delete"
       Exit Sub
    End If
    
    On Error GoTo Del_Err

'Ticket #16189-------------------------------
'    If glbVadim Then
'        DeleteCurrentJob = False
'        SQLQ = "Select TW_JOB,TW_EMPNBR,TW_ID FROM HR_JOB_HISTORY"
'        SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " "
'        SQLQ = SQLQ & " AND TW_CURRENT <>0 "
'        SQLQ = SQLQ & " AND TW_ID = " & Data1.Recordset!TW_ID
'        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'        If Not rsTemp.EOF Then
'            DeleteCurrentJob = True
'        End If
'        rsTemp.Close
'        If DeleteCurrentJob Then
'            If glbMulti Then
'                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID & " AND ED_PAYROLL_ID='" & txtPayrollID & "'"
'                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'                If Not rsTemp.EOF Then
'                    MsgBox "The Current Position can not be deleted. Please enter the End Date instead"
'                    rsTemp.Close
'                    Exit Sub
'                End If
'                rsTemp.Close
'            Else
'                ODHRS = Val(medHours(0))
'                OJOB = clpJob
'            End If
'
'        End If
'    Else
        oJob = clpJob
'    End If
'Ticket #16189-------------------------------

    ODOA = dlpStartDate

    Msg = "Are You Sure You Want To Delete This Record? "
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
        
    Screen.MousePointer = HOURGLASS
    
    'Call procedure to delete the required courses of this position
    'Only if the position is current or tracked for course renewal
    If chkTrackCrsRenewal Or chkCurrent(0) Then
        If chkCurrent(0) Then
            Call Track_Courses_Renewal_Update("Delete", "T")
        Else
            Call Track_Courses_Renewal_Update("Delete", "P")
        End If
    End If

    If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
        If Data1.Recordset("TW_CURRENT") <> 0 Then
            fOldFTE = Data1.Recordset("TW_FTENUM")
        Else
            fOldFTE = 0
        End If
    End If

'Ticket #16189-------------------------------
'    If glbMulti Then
'        oENDDATE = dlpENDDATE.Text
'        If oENDDATE <> "" Then
'            If Not updFollow("D") Then
'                Exit Sub
'            End If
'        End If
'    End If
'Ticket #16189-------------------------------

    xID = Data1.Recordset("TW_ID")
    If chkCurrent(0) Then DeleteCurrentJob = True
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_TEMP_WORK WHERE TW_ID=" & xID
    gdbAdoIhr001.CommitTrans
    
'Ticket #16189-------------------------------
'    'George Jan 26,2006
'    If gsAttachment_DB Then
'        gdbAdoIhr001_DOC.BeginTrans
'        gdbAdoIhr001_DOC.Execute "Delete from HRDOC_JOB_HISTORY where DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR = " & glbLEE_ID & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
'        gdbAdoIhr001_DOC.CommitTrans
'    End If
'    'George Jan 26,2006
'Ticket #16189-------------------------------
    
    Data1.Refresh
    
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        Call Set_Current_Flag
        Call Display_Value
    Else
        Call Display_Value
    End If
    
'Ticket #16189-------------------------------
'    If glbVadim And DeleteCurrentJob Then
'        If glbMulti Then
'            Call DeletePayrollEmp(Date, glbLEE_ID, txtPayrollID.Text)
'        Else
'            UpdateAudit = False
'
'            Dim HRChangs As New Collection
'            If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
'                If isChanged_Field(HRChangs, ODHRS, Data1.Recordset("TW_DHRS"), True) Then UpdateAudit = True
'                If isChanged_Field(HRChangs, OJOB, Data1.Recordset("TW_JOB")) Then UpdateAudit = True
'            Else
'                If isChanged_Field(HRChangs, ODHRS, medHours(0), True) Then UpdateAudit = True
'                If isChanged_Field(HRChangs, OJOB, clpJob) Then UpdateAudit = True
'            End If
'
'            If UpdateAudit = True Then
'                Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID)
'            End If
'        End If
'    End If
'Ticket #16189-------------------------------

    fglbNew = False
    Call SET_UP_MODE

'Ticket #16189-------------------------------
'    If glbGuelph And (Not glbtermopen) Then
'        Call AddFTE(glbLEE_ID, "DELE")
'    End If
'Ticket #16189-------------------------------

'Ticket #16189-------------------------------
'    If glbAdv Then 'Ticket #15282
'        Call Employee_PositionDel_Integration(glbLEE_ID, OJOB, ODOA, True)
'    End If
'Ticket #16189-------------------------------
    
    DoEvents
    
    'Track Courses for the previous Position which turned into Current
    If chkCurrent(0) And DeleteCurrentJob Then
        'Changed 2
        'Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        Call Update_Employee_Job_Training_List(clpJob.Text, "Temporary")
    ElseIf chkTrackCrsRenewal And DeleteCurrentJob Then
        xCurJob = Get_Current_Primary_Job
        If xCurJob <> "" Then
            Call frmEPOSITION.Update_Employee_Job_Training_List(xCurJob, "Current")
        End If
    End If

    Screen.MousePointer = DEFAULT
    
Exit Sub
    
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_TEMP_WORK", "Delete")
    Call RollBack '26July99 js
End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub

Public Sub cmdModify_Click()
    Dim Response%, Msg$, Title$, DgDef As Double
    Dim x% 'jaddy 10/25/99
    
    On Error GoTo Mod_Err
    
    'Ticket #16189-------------------------------
'    If glbGuelph Then
'        medFTENum.Enabled = False
'        medFTEHrs.Enabled = False
'    End If
    'Ticket #16189-------------------------------
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        If chkCurrent(0).Value = True Then
            chkTrackCrsRenewal.Visible = False
        Else
            chkTrackCrsRenewal.Visible = True
        End If
    End If
    
    oTrkCrsRen = chkTrackCrsRenewal
    
    SavFte = medFTENum
    SavFteHr = medFTEHrs
    savCurrent = chkCurrent(0)
    oENDDATE = dlpENDDATE.Text
    oEndReason = clpCode(2).Text
    
    For x% = 0 To 2
        SavRpta(x%) = elpReptAuthShow(x%).Text
    Next
    
    'savWHRS = medHours(1)
    savSDate = dlpStartDate.Text
    savJOB = clpJob.Text
    fglbNew% = False
    glbChgTermDate = ""
    glbChgTermReason = ""
    
    'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
'        OBillingRate = medBillingRate
'    End If
    'Ticket #16189-------------------------------
    
Exit Sub
    
Mod_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_TEMP_WORK", "Modify")
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
    
    If Not Data1.Recordset.EOF Then
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
        
        If (chkTrackCrsRenewal.Visible = True And chkTrackCrsRenewal.Value = False) Or chkCurrent(0) Then   'New Tracking method
            'Confirm the Tracking Course Renewal ON or OFF
            Msg = "Do you want to track required courses renewals for this position? "
            xRes = MsgBox(Msg, vbYesNoCancel, "Confirm Required Course Renewal Tracking")
            If xRes = 7 Then    'No
                'If chkCurrent(0) Then   'New Tracking method    'Ticket #22951
                    'Delete required courses
                    Call Track_Courses_Renewal_Update("Delete", "T")
                'End If
                
                'Update records with tracking on
                chkTrackCrsRenewal.Value = False
                rsDATA("TW_TRK_CRS_RENEWAL") = False
                rsDATA.Update
                
                GoTo Continue_NewClick
            ElseIf xRes = 2 Then    'Cancel
                Exit Sub
            Else
                'Delete required courses ANYWAYS so that it can be added correctly with right type of position
                'Changed
                Call Track_Courses_Renewal_Update("Delete", "T")
            End If
            
            'Turn-ON tracking
            chkTrackCrsRenewal.Visible = True
            chkTrackCrsRenewal.Value = True
            
            'Call procedure to update/delete employee's Training list with this position's
            'required course list
            Call Track_Courses_Renewal_Update
            
            'Update records with tracking on
            rsDATA("TW_TRK_CRS_RENEWAL") = True
            rsDATA.Update
            
            chkTrackCrsRenewal.Value = False
            'chkTrackCrsRenewal.Visible = False
            DoEvents
        End If
    End If
    
Continue_NewClick:
    fglbNew = True
    flgNewCancel = True
    
    If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
        fFTEDate = CurSDate
        fOldFTE = OFTE
    End If
    
    If glbLinamar Then
        clpJob.TransDiv = Right(glbLEE_ID, 3)
    End If
    
'Ticket #16189-------------------------------
'    If glbLambton Then
'        chkUseForBenefit.Visible = True
'    End If
'Ticket #16189-------------------------------

    Action = "A"
    SavFte = ""
    SavFteHr = ""
    txtLambtonJob = ""
    fglbNew% = True
    
    Call SET_UP_MODE
    
'Ticket #16189-------------------------------
'    'George on Jan 26,2006 #10266
'    If gsAttachment_DB Then
'        glbJob = ""
'        glbSDate = "01/01/1900"
'        lblImport.Visible = True
'        imgSec.Visible = False
'        imgNoSec.Visible = True
'        cmdImport.Visible = True
'    End If
'    'George on Jan 26,2006 #10266
'Ticket #16189-------------------------------

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
    End If
    If glbCompSerial = "S/N - 2379W" Then 'Town of LaSalle Ticket #14534
        txtShift.Text = "NOSD"
    End If
    
    rsDATA.AddNew
        
    chkCurrent(0) = glbMulti
    lblEEID = glbLEE_ID
    lblCompNo.Caption = "001"
    Call SetDefaultValue
    
'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2241W" Then 'Granite Club
'        If NewHireForms.count > 0 Then 'New Hire only
'            chkActPosition.Value = True
'        End If
'    End If
'Ticket #16189-------------------------------

'Ticket #16189-------------------------------
'    If NewHireForms.count > 0 Then 'From v7.6
'        dlpStartDate = GetDoh(glbLEE_ID)
'    End If
'Ticket #16189-------------------------------

    clpJob.Enabled = True
    clpJob.SetFocus
    
    'Simona - begin - Assessment Strategies-#14963
    If (glbCompSerial = "S/N - 2401W") Then
        medHours(0).Text = "7.5"
        medHours(1).Text = "37.5"
        medHours(2).Text = "75.0"
        medFTENum.Text = "1"
        medFTEHrs.Text = "1950"
    End If
    'Simona - end - Assessment Strategies-#14963

    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        cmdJobFiles.Enabled = False
        chkTrackCrsRenewal.Enabled = False
        chkTrackCrsRenewal.Visible = False
    End If

Exit Sub

AddNP_Err:
    If Err = 3018 Then
        Err = 0
        Resume Next
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_TEMP_WORK", "Add")
    Call RollBack '26July99 js
End Sub

Public Sub cmdOK_Click()
Dim x%, xID, xFte, xFteHr
Dim rsJOB As New ADODB.Recordset
Dim rsJOBMASTER As New ADODB.Recordset
Dim xReptAuthority, xChange
Dim SQLQ, Msg, startDate
Dim rs As New ADODB.Recordset
Dim oCurrent As Boolean

On Error GoTo Add_Err

'Ticket #16189-------------------------------
''City of Timmins - Ticket #13207
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
'Ticket #16189-------------------------------

'Ticket #16189-------------------------------
'If glbVadim And glbMulti And Not fglbNew Then
'    If Not chkVadimPayrollID Then Exit Sub
'End If
'Ticket #16189-------------------------------

If Not chkPosition() Then Exit Sub

Dim xRes As Integer

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
            If Not fglbNew Then
                Call Track_Courses_Renewal_Update
            End If
        End If
        
        'Hold current value of the Current flag
        oCurrent = chkCurrent(0).Value
    Else
        'Hold current value of the Current flag
        oCurrent = chkCurrent(0).Value
    End If
Else
    'Hold current value of the Current flag
    oCurrent = chkCurrent(0).Value
End If

'Ticket #16189-------------------------------
'If Not glbSetPos Then Call UpdPositionCCAC
'Ticket #16189-------------------------------
Screen.MousePointer = HOURGLASS

Call UpdUStats(Me) ' update user's stats (who did it and when)

If chkTrackCrsRenewal.Value = False Or chkTrackCrsRenewal.Visible = False Then
    chkTrackCrsRenewal.Value = False
    rsDATA("TW_TRK_CRS_RENEWAL") = False
End If

'Ticket #16189-------------------------------
'If glbCompSerial = "S/N - 2259W" And (Not glbtermopen) Then
'    SQLQ = "SELECT ED_ORG, ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
'    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'    If rs.EOF = False And rs.BOF = False Then
'        If rs("ED_SECTION") <> "Y" Then
'            If IsNull(rs("ED_ORG")) = False Then clpCode(0).Text = rs("ED_ORG")
'        End If
'    End If
'    rs.Close
'    Set rs = Nothing
'End If
'Ticket #16189-------------------------------
'Ticket #16189-------------------------------
''City of Pickering - Ticket #13281
'If glbCompSerial = "S/N - 2217W" Then
'    If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
'        If IsNumeric(medHours(2)) And (medFTEHrs.Text = "" Or medFTEHrs.Text = "0") Then     'Hours/Pay Period
'            medFTEHrs = medHours(2) * 26
'        End If
'    End If
'End If
'If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
'    If fglbNew Then
'        If IsNumeric(medFTENum) Then
'            fNewFTE = Val(medFTENum)
'        Else
'            fNewFTE = 0
'        End If
'    End If
'End If
'Ticket #16189-------------------------------

Call Set_Control("U", Me, rsDATA)

For x = 0 To 2
    xReptAuthority = getEmpnbr(elpReptAuthShow(x))
    rsDATA("TW_REPTAU" & IIf(x = 0, "", x + 1)) = IIf(Val(xReptAuthority) = 0, Null, xReptAuthority)
Next
If glbLinamar Then
    'If chkActPosition Then
    '    rsDATA("TW_POSITION_CONTROL") = "YES"
    'Else
    '    rsDATA("TW_POSITION_CONTROL") = "NO"
    'End If
ElseIf glbMulti Then 'George on Dec 7,2005 #9928 begin
    If chkActPosition Then
        If fglbNew Then
            xID = 0
        Else
            xID = rsDATA!TW_ID
        End If
        SQLQ = "UPDATE HR_TEMP_WORK"
        SQLQ = SQLQ & " SET TW_POSITION_CONTROL = 'NO' "
        SQLQ = SQLQ & " WHERE TW_EMPNBR =" & glbLEE_ID & " AND TW_ID <> " & xID
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
        'rsDATA("TW_POSITION_CONTROL") = "YES"
    Else
        'rsDATA("TW_POSITION_CONTROL") = "NO"
    End If 'George on Dec 7,2005 #9928 end
End If

'Ticket #16189-------------------------------
'Hemu - Ottawa CCAC uses this field to record their own CCAC Position # and so it cannot be
'set to NO or YES - Ticket #11411
'If Not glbOttawaCCAC Then
'    If chkActPosition Then
'        rsDATA("TW_POSITION_CONTROL") = "YES"
'    Else
'        rsDATA("TW_POSITION_CONTROL") = "NO"
'    End If
'End If
'Ticket #16189-------------------------------

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    'gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    xID = rsDATA!TW_ID
    'gdbAdoIhr001X.CommitTrans
    'rsDATA.Resync
    
    'Ticket #16189-------------------------------
'    'George Jan 26,2006
'    If gsAttachment_DB Then
'        gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_JOB_HISTORY set DJ_JOB='" & rsDATA("TW_JOB") & "',DJ_SDATE=" & Date_SQL(rsDATA("TW_SDATE")) & " where DJ_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
'    End If
'    'George Jan 26,2006
    'Ticket #16189-------------------------------
Else
    'gdbAdoIhr001.BeginTrans
    rsDATA.Update
    xID = rsDATA!TW_ID
    'gdbAdoIhr001.CommitTrans
    'rsDATA.Requery
    'xID = rsDATA!TW_ID
    
    'Ticket #16189-------------------------------
'    'George Jan 26,2006
'    If gsAttachment_DB Then
'        gdbAdoIhr001_DOC.Execute "Update HRDOC_JOB_HISTORY set DJ_JOB='" & rsDATA("TW_JOB") & "',DJ_SDATE=" & Date_SQL(rsDATA("TW_SDATE")) & " where DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR = " & glbLEE_ID & " and DJ_JOB='" & glbJob & "' and DJ_SDATE=" & Date_SQL(glbSDate)
'    End If
'    'George Jan 26,2006
    'Ticket #16189-------------------------------
End If

'Ticket #16189-------------------------------
''Add by Franks on Jul 11,02 for ticket #2546
''If glbWFC And lblBANDCode.Visible Then
'If glbWFC And clpCode(6).Visible Then
'    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & clpJob.Text & "' "
'    rsJOBMASTER.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsJOBMASTER.EOF Then
'        If rsJOBMASTER("JB_BAND") <> clpCode(6).Text Then
'            rsJOBMASTER("JB_BAND") = clpCode(6).Text
'            rsJOBMASTER.Update
'            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_BAND = '" & clpCode(6).Text & "' "
'            SQLQ = SQLQ & "WHERE SH_JOB = '" & clpJob.Text & "' "
'            gdbAdoIhr001.BeginTrans
'            gdbAdoIhr001.Execute SQLQ
'            gdbAdoIhr001.CommitTrans
'        End If
'    End If
'    rsJOBMASTER.Close
'End If
''Add by Franks on Jul 11,02 for ticket #2546
'Ticket #16189-------------------------------

'Data1.Refresh

'Ticket #16189-------------------------------
'Burlington Tech Ticket #13235
''If the new position code is found in Backup Position table, delete it.
'If glbCompSerial = "S/N - 2351W" Then
'    If fglbNew And (Not glbtermopen) Then
'        SQLQ = "DELETE FROM HR_JOB_BACKUP WHERE TW_EMPNBR = " & glbLEE_ID & " "
'        SQLQ = SQLQ & "AND TW_JOB = '" & clpJob.Text & "' "
'        gdbAdoIhr001.Execute SQLQ
'    End If
'End If
'Ticket #16189-------------------------------

Call Set_Current_Flag
Data1.Refresh
Data1.Recordset.Find "TW_ID=" & xID

'Ticket #16189-------------------------------
'If gsAttachment_DB Then
'    If glbDocNewRecord Then 'New Record only
'        If Len(glbDocImpFile) > 0 Then
'            'glbJob = xID
'            glbJob = Data1.Recordset("TW_JOB")
'            glbSDate = Data1.Recordset("TW_SDATE")
'            Call AttachmentAdd(glbLEE_ID, glbDocImpFile)
'        End If
'    End If
'    glbDocImpFile = ""
'End If
'Ticket #16189-------------------------------

chkCurrent(0) = Data1.Recordset("TW_CURRENT")

'Call procedure to add required courses to the Training List
If fglbNew And chkCurrent(0) Then
    'Changed 2
    'Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
    Call Update_Employee_Job_Training_List(clpJob.Text, "Temporary")
ElseIf fglbNew And Not chkCurrent(0) Then
    Call Track_Courses_Renewal_Update
Else
    'Ticket #22951
    'Position Code has changed. Delete the Training List of the older Position and then create new
    'Training List for then changed Position Code. (Ticket #22044)
    If Len(savJOB) > 0 And Not savJOB = clpJob.Text And chkCurrent(0) Then
        Call Track_Courses_Renewal_Update("Delete", "T", savJOB)
        Call Update_Employee_Job_Training_List(clpJob.Text, "Temporary")
    End If

    If (chkCurrent(0) And oCurrent <> chkCurrent(0)) Or (chkCurrent(0) And savCurrent <> chkCurrent(0)) Then    'New Tracking method
        'Changed 2
        'Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        Call Update_Employee_Job_Training_List(clpJob.Text, "Temporary")
    ElseIf Not fglbNew And Not chkCurrent(0) And chkTrackCrsRenewal.Value = False Then  'Ticket #22951
        Call Track_Courses_Renewal_Update("Delete", "T")
    End If
End If

'Ticket #16189-------------------------------
'If chkCurrent(0) Then
'    SQLQ = ""
'    For x = 0 To 2
'        xReptAuthority = getEmpnbr(elpReptAuthShow(x))
'        If SavRpta(x%) <> elpReptAuthShow(x%).Text Then xChange = True
'        SQLQ = SQLQ & " PH_REPTAU" & IIf(x = 0, "", x + 1) & " =" & IIf(Val(xReptAuthority) > 0, xReptAuthority, "Null") & IIf(x = 2, " ", ",")
'    Next
'    If Action = "M" Then
'        If savJOB <> clpJob.Text Or savSDate <> dlpStartDate.Text Then
'        Else
'            If xChange Then
'                SQLQ = "UPDATE HR_PERFORM_HISTORY SET " & SQLQ
'                SQLQ = SQLQ & " WHERE PH_EMPNBR=" & glbLEE_ID & " AND PH_JOB='" & clpJob.Text & "' AND PH_CURRENT<>0 "
'                gdbAdoIhr001.Execute SQLQ
'            End If
'        End If
'        If savWHRS <> medHours(1) Then
'            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_WHRS =" & Val(medHours(1))
'            SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_JOB='" & clpJob.Text & "' AND SH_CURRENT<>0 "
'            gdbAdoIhr001.Execute SQLQ
'
'            'Hemu
'            savWHRS = medHours(1)
'            'Hemu
'
'        End If
'        If savGrid <> clpGrid.Text And glbMultiGrid Then
'            SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_GRID ='" & clpGrid.Text & "'"
'            SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_JOB='" & clpJob.Text & "' AND SH_CURRENT<>0 "
'            gdbAdoIhr001.Execute SQLQ
'            savGrid = clpGrid.Text
'        End If
'
'        'End If
'    End If
'    If Not glbMulti Then
'        SQLQ = "UPDATE HREMP SET ED_SHIFT ='" & txtShift & "' WHERE ED_EMPNBR=" & glbLEE_ID
'        gdbAdoIhr001.Execute SQLQ
'    End If
'    If Val(medHours(0)) <> Val(fgtxtDhrs) Then
'        SQLQ = "UPDATE HREMP SET ED_DHRS =" & Val(medHours(0)) & " WHERE ED_EMPNBR=" & glbLEE_ID
'        gdbAdoIhr001.Execute SQLQ
'        SQLQ = "UPDATE HRENTHRS SET HE_DHRS =" & Val(medHours(0)) & " WHERE HE_EMPNBR=" & glbLEE_ID
'        gdbAdoIhr001.Execute SQLQ
'        glbENTScreen = True
'    End If
'    If SavFte <> medFTENum Or SavFteHr <> medFTEHrs Then
'        If SavFte <> medFTENum Then xFte = SavFte Else xFte = ""
'        If SavFteHr <> medFTEHrs Then xFteHr = SavFteHr Else xFteHr = ""
'        If Not EmpHisCalc(3, glbLEE_ID, "", "", "", "", "", xFte, xFteHr, Date) Then MsgBox "EMPHIS Error"
'    End If
'    If Not glbMulti Then
'        'Hemu 07/02/2003 Begin - Ticket #4247, Update Employment Equity Data with NOC Code
'        Dim rsEmpNOC As New ADODB.Recordset
'
'        rsEmpNOC.Open "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT TW_JOB FROM HR_JOB_HISTORY WHERE TW_EMPNBR = " & glbLEE_ID & " AND TW_CURRENT <> 0)", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        If Not rsEmpNOC.EOF Then
'            If Not IsNull(rsEmpNOC("JB_FEDGRP")) Then
'                gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NOGC = '" & rsEmpNOC("JB_FEDGRP") & "' WHERE EQ_EMPNBR = " & glbLEE_ID
'            End If
'        End If
'        rsEmpNOC.Close
'        'Hemu 07/02/2003 End - Ticket #4247
'    End If
'
'    Call InitData
'End If
'Ticket #16189-------------------------------

'Ticket #16189-------------------------------
'If glbCompSerial = "S/N - 2217W" Then ' FOR CITY OF PICKERING
'    If chkCurrent(0) Then
'        If Not updFollow("U") Then Exit Sub
'    End If
'End If
'
'If glbGuelph And (Not glbtermopen) Then  ' FOR Guelph-Willington
'    If fglbNew Then
'        Call Pause(0.5)
'        Call AddFTE(glbLEE_ID, "NEW")
'    End If
'End If
'If glbOttawaCCAC Then
'   If chkCurrent(0) Then
'        Call UpdOttawaCCAC
'   End If
'End If
'If glbCompSerial = "S/N - 2347W" Then
'    Call updBenefitForSurreyPlace(glbLEE_ID)
'End If
'If Not glbMediPay Then
'    Call Employee_Master_Integration(glbLEE_ID)
'End If
'Ticket #16189-------------------------------

'George Mar 9 2006 commented. Moved to Upd_Related_Salary. Here could not know this position changed will create a new salary record in Salary_History or just change.
'If glbCompSerial = "S/N - 2259W" Or glbGP Then
'    Call Salary_Integration(glbLEE_ID, , False, IIf(fglbNew% = 0, False, True))
'End If
'aded by Bryan 22/09/05 Ticket# 9368

'Ticket #16189-------------------------------
'If glbMediPay Then 'Ticket #14752
'    'Hemu - Ticket #14752 - Because Job Start Date and Reason for Change needs to be passed
'    'as well as Salary Effective Date and Reason for Change whenever these happens, I had to
'    'pass to separate the function out.
'    'Call Salary_Integration(glbLEE_ID)
'    Call Position_Integration(glbLEE_ID)
'End If
'
'If NewHireForms.count > 0 And glbCompSerial = "S/N - 2375W" Then
'    Call updateOMERS
'End If

''added by Bryan 12/Apr/06 Ticket#10644
'If isEDU Then
'    If elpReptAuthShow(0).Text <> "" Then
'        If glbCompSerial = "S/N - 2347W" And NewHireForms.count > 0 Then 'Surreyplace
'
'
'            SQLQ = "SELECT HRE_SCHEDULE.SC_CLASSID, HRE_SCHEDULE.SC_DATE FROM HRE_COURSE INNER JOIN HRE_SCHEDULE ON HRE_COURSE.CS_ID = HRE_SCHEDULE.SC_CLASSID "
'            SQLQ = SQLQ & "WHERE HRE_SCHEDULE.SC_DATE > " & Date_SQL(Date) & " AND HRE_COURSE.CS_CODE='ORIE' ORDER BY HRE_SCHEDULE.SC_DATE ASC"
'            rs.Open SQLQ, gdbAdoIHREDU, adOpenStatic, adLockOptimistic, adCmdText
'            If rs.EOF = False And rs.BOF = False Then
'                SQLQ = "INSERT INTO HRE_ENROLLMENT(EN_EMPNBR, EN_TYPE, EN_CLASSID, EN_WAITING, EN_NAME, EN_SUPER, EN_LDATE, EN_LUSER, EN_LTIME) "
'                SQLQ = SQLQ & "VALUES (" & glbLEE_ID & ", 'E', " & rs("SC_CLASSID") & ", 1, '" & Replace(lblEEName.Caption, "'", "") & "'," & elpReptAuthShow(0).Text & ", "
'                SQLQ = SQLQ & Updstats(0) & ", '" & Updstats(2) & "', '" & Updstats(1) & "')"
'                gdbAdoIHREDU.BeginTrans
'                gdbAdoIHREDU.Execute SQLQ
'                gdbAdoIHREDU.CommitTrans
'            End If
'            rs.Close
'        End If
'
'        If glbCompSerial = "S/N - 2347W" And fglbNew Then 'Surreyplace
'            Dim xclass As String
'
'            SQLQ = "SELECT HRE_SCHEDULE.SC_CLASSID, HRE_COURSE.CS_CODE, HRE_SCHEDULE.SC_DATE FROM HRE_COURSE INNER JOIN HRE_SCHEDULE ON HRE_COURSE.CS_ID = HRE_SCHEDULE.SC_CLASSID "
'            SQLQ = SQLQ & "WHERE HRE_SCHEDULE.SC_DATE > " & Date_SQL(Date) & " AND HRE_COURSE.CS_CODE IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB='" & clpJob.Text & "')"
'            SQLQ = SQLQ & "ORDER BY HRE_SCHEDULE.SC_DATE ASC"
'            rs.Open SQLQ, gdbAdoIHREDU, adOpenStatic, adLockOptimistic, adCmdText
'            If rs.EOF = False And rs.BOF = False Then
'                xclass = ""
'                Do
'                    If xclass <> rs("CS_CODE") Then
'                        xclass = rs("CS_CODE")
'                        SQLQ = "INSERT INTO HRE_ENROLLMENT(EN_EMPNBR, EN_TYPE, EN_CLASSID, EN_WAITING, EN_NAME, EN_SUPER, EN_LDATE, EN_LUSER, EN_LTIME) "
'                        SQLQ = SQLQ & "VALUES (" & glbLEE_ID & ", 'E', " & rs("SC_CLASSID") & ", 1, '" & Replace(lblEEName.Caption, "'", "") & "'," & elpReptAuthShow(0).Text & ", "
'                        SQLQ = SQLQ & Updstats(0) & ", '" & Updstats(2) & "', '" & Updstats(1) & "')"
'                        gdbAdoIHREDU.BeginTrans
'                        gdbAdoIHREDU.Execute SQLQ
'                        gdbAdoIHREDU.CommitTrans
'                    End If
'                    rs.MoveNext
'                Loop Until rs.EOF
'            End If
'            rs.Close
'        End If
'    End If
'Set rs = Nothing
'End If
''end Bryan
'Ticket #16189-------------------------------

ExitLine1:

fglbNew = False
flgNewCancel = False
Call Display_Value

Screen.MousePointer = DEFAULT
'Ticket #16189-------------------------------
'If NewHireForms.count > 0 Then
'    glbLinNewPosSal = True
'End If
'
'If glbLinamar And Action = "A" And NewHireForms.count = 0 Then
'    Msg = "Do you want update the employee's Payroll and Personnel information? "
'    If MsgBox(Msg, 36, "info:HR") = 6 Then
'     frmBasicLinamar.Show 1
'    End If
'End If
'If glbCompSerial = "S/N - 2291W" And Action = "A" And NewHireForms.count = 0 Then
'    Msg = "Do you want update the employee's demographics? "
'    If MsgBox(Msg, 36, "info:HR") = 6 Then
'        frmBasicSyndesis.Show 1
'    End If
'End If
'If glbOttawaCCAC Then
'    If chkCurrent(0) Then
'        If IsNull(GetSHData(glbLEE_ID, "SH_PAYP", Null)) Then
'            If medHours(1) = 0 And medHours(2) = 0 Then
'                MsgBox "Please enter Hours/Week and Hours/Pay Period if this is an ""E""-""Exceptional Hourly"" employee."
'                medHours(1).SetFocus
'                Exit Sub
'            Else
'                If medHours(1) = 0 Then
'                    MsgBox "Please enter Hours/Week if this is an ""E""-""Exceptional Hourly"" employee."
'                    medHours(1).SetFocus
'                    Exit Sub
'                End If
'                If medHours(2) = 0 Then
'                    MsgBox "Please enter Hours/Pay Period if this is an ""E""-""Exceptional Hourly"" employee."
'                    medHours(2).SetFocus
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
'End If
'Ticket #16189-------------------------------

Action = "M"

'Ticket #16189-------------------------------
'Call NextForm
'Ticket #16189-------------------------------

Exit Sub

Add_Err:
    If Err = 3018 Then
        Err = 0
        Resume Next
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_TEMP_WORK", "Update")
    Call RollBack '26July99 js
    Resume Next
End Sub

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
    RHeading = lblEEName & "'s Temporary/Cross Training Position History"
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
    RHeading = lblEEName & "'s Temporary/Cross Training Position History"
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
Dim HRTW_Snap As New ADODB.Recordset

CurSDate = 0    ' returns 0 if no found records

On Error GoTo JHS_Err

SQLQ = "Select HR_TEMP_WORK.* FROM HR_TEMP_WORK"
SQLQ = SQLQ & " WHERE HR_TEMP_WORK.TW_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND HR_TEMP_WORK.TW_CURRENT <>0"

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
HRTW_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRTW_Snap.BOF And HRTW_Snap.EOF Then
    Exit Function
Else
    'Ticket #16189-------------------------------
'    'Not (Town of Aurora and City of Timmins and City of Kawartha Lakes)
'    If glbVadim And glbMulti And glbCompSerial <> "S/N - 2378W" And glbCompSerial <> "S/N - 2375W" And glbCompSerial <> "S/N - 2363W" Then
'        If fglbNew Then
'            If empPayrollID = txtPayrollID Then
'                SetEmpValue (True)
'            End If
'            Do While Not HRTW_Snap.EOF
'                If HRTW_Snap("TW_PAYROLL_ID") = txtPayrollID Then
'                    CurSDate = HRTW_Snap("TW_SDATE")
'                    oPHRS = HRTW_Snap("TW_PHRS")
'                    ODHRS = HRTW_Snap("TW_DHRS")
'                    oWHRS = HRTW_Snap("TW_WHRS")
'                    OJOB = HRTW_Snap("TW_JOB")
'                    OSDATE = HRTW_Snap("TW_SDATE")
'                    OReason = HRTW_Snap("TW_JREASON")
'                    oPayrollID = HRTW_Snap("TW_PAYROLL_ID")
'                    oOrg = HRTW_Snap("TW_ORG")
'                    oDeptNo = HRTW_Snap("TW_DEPTNO")
'                    oGLNo = HRTW_Snap("TW_GLNO")
'                    oStatus = HRTW_Snap("TW_EMP")
'                    oPayCategory = HRTW_Snap("TW_PAYROLL_CATEGORY")
'                    If IsNull(HRTW_Snap("TW_SHIFT")) Then
'                        oSHIFT = ""
'                    Else
'                        oSHIFT = HRTW_Snap("TW_SHIFT")
'                    End If
'                    If IsNull(HRTW_Snap("TW_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRTW_Snap("TW_REPTAU")
'                    If IsNull(HRTW_Snap("TW_GRID")) Then
'                        OLambtonJob = OJOB
'                    Else
'                        OLambtonJob = Left(HRTW_Snap("TW_GRID"), 1) & OJOB & Mid(HRTW_Snap("TW_GRID"), 2)
'                    End If
'                    HRTW_Snap("TW_CURRENT") = 0
'                    HRTW_Snap.Update
'                End If
'                HRTW_Snap.MoveNext
'            Loop
'
'            HRTW_Snap.Close
'        Else
'            CurSDate = 0
'            If empPayrollID = txtPayrollID Then
'                SetEmpValue
'            End If
'            oPHRS = Data1.Recordset("TW_PHRS")
'            ODHRS = Data1.Recordset("TW_DHRS")
'            oWHRS = Data1.Recordset("TW_WHRS")
'            OJOB = Data1.Recordset("TW_JOB")
'            OSDATE = Data1.Recordset("TW_SDATE")
'            OReason = Data1.Recordset("TW_JREASON")
'            oPayrollID = Data1.Recordset("TW_PAYROLL_ID")
'            oOrg = Data1.Recordset("TW_ORG")
'            oDeptNo = Data1.Recordset("TW_DEPTNO")
'            oGLNo = Data1.Recordset("TW_GLNO")
'            oStatus = Data1.Recordset("TW_EMP")
'            oPayCategory = Data1.Recordset("TW_PAYROLL_CATEGORY")
'            If IsNull(HRTW_Snap("TW_SHIFT")) Then
'                oSHIFT = ""
'            Else
'                oSHIFT = Data1.Recordset("TW_SHIFT")
'            End If
'            'oREPTAU = Data1.Recordset("TW_REPTAU")
'            If IsNull(Data1.Recordset("TW_REPTAU")) Then oREPTAU = "" Else oREPTAU = Data1.Recordset("TW_REPTAU")
'            If IsNull(Data1.Recordset("TW_GRID")) Then
'                OLambtonJob = OJOB
'            Else
'                OLambtonJob = Left(Data1.Recordset("TW_GRID"), 1) & OJOB & Mid(Data1.Recordset("TW_GRID"), 2)
'            End If
'        End If
'    Else
'Ticket #16189-------------------------------
    If glbMulti Then
        Do While Not HRTW_Snap.EOF
            If HRTW_Snap("TW_JOB") = clpJob.Text Then
                CurSDate = HRTW_Snap("TW_SDATE")
                oPHRS = HRTW_Snap("TW_PHRS")
                ODHRS = HRTW_Snap("TW_DHRS")
                oWHRS = HRTW_Snap("TW_WHRS")
                oJob = HRTW_Snap("TW_JOB")
                OSDATE = HRTW_Snap("TW_SDATE")
                OReason = HRTW_Snap("TW_JREASON")
                oPayrollID = HRTW_Snap("TW_PAYROLL_ID")
                oOrg = HRTW_Snap("TW_ORG")
                oDeptNo = HRTW_Snap("TW_DEPTNO")
                oGLNo = HRTW_Snap("TW_GLNO")
                oStatus = HRTW_Snap("TW_EMP")
                If IsNull(HRTW_Snap("TW_SHIFT")) Then
                    oSHIFT = ""
                Else
                    oSHIFT = HRTW_Snap("TW_SHIFT")
                End If
                If IsNull(HRTW_Snap("TW_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRTW_Snap("TW_REPTAU")
            End If
            HRTW_Snap.MoveNext
        Loop
        HRTW_Snap.Close
    Else
        CurSDate = HRTW_Snap("TW_SDATE")
        oPHRS = HRTW_Snap("TW_PHRS")
        ODHRS = HRTW_Snap("TW_DHRS")
        oWHRS = HRTW_Snap("TW_WHRS")
        oJob = HRTW_Snap("TW_JOB")
        OSDATE = HRTW_Snap("TW_SDATE")
        OReason = HRTW_Snap("TW_JREASON")
        If IsNull(HRTW_Snap("TW_SHIFT")) Then
            oSHIFT = ""
        Else
            oSHIFT = HRTW_Snap("TW_SHIFT")
        End If
        If IsNull(HRTW_Snap("TW_REPTAU")) Then oREPTAU = "" Else oREPTAU = HRTW_Snap("TW_REPTAU")
        If glbLinamar Then
            OLeadHand = HRTW_Snap("TW_LEADHAND")
            OLabourCD = HRTW_Snap("TW_LABOURCD")
            oLABOUREDATE = HRTW_Snap("TW_LABOUREDATE")
        End If
        HRTW_Snap.Close
    End If
End If

Exit Function
JHS_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Temp/Cross Training History Snap", "HR_TEMP_WORK", "SELECT")
    Call RollBack '26July99 js
End Function

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    glbFrmCaption$ = Me.Caption
    glbErrNum& = ErrorNumber
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_TEMP_WORK", "SELECT")
    Call RollBack '26July99 js
End Sub

Public Function EERetrieve()
Dim SQLQ As String
Dim x, xFld
Dim rt As New ADODB.Recordset

EERetrieve = False

On Error GoTo EERError
    
    Screen.MousePointer = HOURGLASS
        
    'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2259W" Then 'Added by Bryan 11/07/05 Ticket #8857
'        If glbtermopen Then
'            SQLQ = "Select ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
'        Else
'            SQLQ = "Select ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
'        End If
'        Dim rs As New ADODB.Recordset
'        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
'        If rs("ED_SECTION") = "Y" Then
'            glbMulti = True
'            frmMulti.Visible = True
'        Else
'            glbMulti = False
'            frmMulti.Visible = False
'        End If
'        rs.Close
'        Set rs = Nothing
'        SQLQ = ""
'    End If
    
'    If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab 'Ticket #14791
'        locOrg = ""
'        SQLQ = "SELECT ED_EMPNBR, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
'        rt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly
'        If Not IsNull(rt("ED_ORG")) Then
'            locOrg = rt("ED_ORG")
'        End If
'        rt.Close
'        Set rt = Nothing
'        SQLQ = ""
'    End If
    'Ticket #16189-------------------------------

    If glbtermopen Then
        SQLQ = "Select Term_TEMP_WORK.*,"
    Else
        SQLQ = "Select HR_TEMP_WORK.*,"
    End If

    For x = 0 To 2
        xFld = "REPTAU" & IIf(x = 0, "", x + 1)
        If glbLinamar Then
            SQLQ = SQLQ & " CASE WHEN TW_" & xFld & " IS NOT NULL AND LEN(TW_" & xFld & ")>2 "
            SQLQ = SQLQ & " THEN RIGHT(TW_" & xFld & ",3)+'-'+"
            SQLQ = SQLQ & " LEFT(TW_" & xFld & ",LEN(TW_" & xFld & ")-3) "
            SQLQ = SQLQ & " ELSE STR(TW_" & xFld & ") END "
            SQLQ = SQLQ & " AS " & xFld & IIf(x = 2, "", ",")
        Else
            If glbOracle Then
                SQLQ = SQLQ & "TW_" & xFld & " AS " & xFld & IIf(x = 2, "", ",")
            Else
                SQLQ = SQLQ & "STR(TW_" & xFld & ") AS " & xFld & IIf(x = 2, "", ",")
            End If
        End If
    Next
    
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_TEMP_WORK"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = SQLQ & " FROM HR_TEMP_WORK"
        SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY "
    
    If glbMulti Then SQLQ = SQLQ & "TW_CURRENT " & IIf(glbSQL, "DESC", "") & ","
    SQLQ = SQLQ & "TW_SDATE DESC"
    
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

    'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
'        lblHrsPayPeriod.FontBold = True
'    ElseIf glbCompSerial = "S/N - 2357W" And glbEmpCountry <> "CANADA" Then   'I.T. Xchange
'        lblHrsPayPeriod.FontBold = False
'    End If
    'Ticket #16189-------------------------------
    
    Screen.MousePointer = DEFAULT
    EERetrieve = True

Exit Function
EERError:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_TEMP_WORK", "SELECT")
    Call RollBack '26July99 js
End Function

Private Sub Form_Activate()
    glbOnTop = "FRMETMPCRSTRNPOS"
    flgloaded = True
    Call Job_Desc
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMETMPCRSTRNPOS"
End Sub

Sub Form_Load()
    Dim Answer, DefVal, Msg, Title  '  variables.
    Dim RFound As Integer ' records found
    Dim x%
    Dim rsTA As New ADODB.Recordset
    
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    glbOnTop = "FRMETMPCRSTRNPOS"
        
    If glbtermopen Then
        Data1.ConnectionString = glbAdoIHRAUDIT
    Else
        Data1.ConnectionString = glbAdoIHRDB
    End If
    
    If glbMulti Or glbLinamar Then
        frmMulti.Visible = True
        If glbLinamar Then 'Ticket# 8293
            frmMulti.Height = 440
        End If
    End If
    
    frmJobEnd.BorderStyle = 0
    
    'Ticket #16189-------------------------------
'    If glbWFC Then 'Ticket #14927
'        frmJobEnd.Visible = False
'        'Ticket #15396 - begin
'        lblReptAuth(0).FontBold = True
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'        'Ticket #15396 - end
'    End If
    
'    lblBand.Visible = glbWFC
'    clpCode(6).Visible = glbWFC
    
'    If glbWFC Then 'Ticket #11772
'        cboShift.Left = txtShift.Left
'        cboShift.Visible = True
'        txtShift.Visible = False
'        cboShift.AddItem "NS"
'        cboShift.AddItem "1"
'        cboShift.AddItem "2"
'        cboShift.AddItem "3"
'        cboShift.AddItem "4"
'        cboShift.AddItem "5"
'        cboShift.AddItem "6"
'        cboShift.AddItem "A"
'        cboShift.AddItem "B"
'        cboShift.AddItem "C"
'        cboShift.AddItem "D"
'        cboShift.AddItem "E"
'        cboShift.AddItem "F"
'        cboShift.AddItem "M"
'        cboShift.AddItem "Q"
'        cboShift.AddItem "R"
'        cboShift.AddItem "S"
'        cboShift.AddItem "T"
'        cboShift.AddItem "W"
'    End If
    'Ticket #16189-------------------------------

    If glbMultiGrid Then
        lblGrid.Visible = True
        clpGrid.Visible = True
    Else
        lblPosTitle.Top = lblGrid.Top
        clpJob.Top = clpGrid.Top
    End If
    Call CR_Job_Snap
    Screen.MousePointer = HOURGLASS
    
    glbLinNewPosSal = False
    lblSection.Caption = "Section" 'St. John's relabelled to Section - NA. So on scrolling thru one emply to another - the program is duplicating the label within itself.
    
    Call setCaption(lblGrid)
    
    Screen.MousePointer = DEFAULT
    
    'Ticket #16189-------------------------------
'    If glbLinamar Then
'        For x = 0 To 4
'            frmLinamar(x).Visible = True
'        Next
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'        lblShift.FontBold = True
'    Else
    'Ticket #16189-------------------------------
        If glbMulti Then 'George on Dec 7,2005 #9928 begin
            chkActPosition.Caption = "Default Position"
        End If 'George on Dec 7,2005 #9928 end
        panControls.Height = 0
'Ticket #16189-------------------------------
'    End If
    
'    If glbVadim Then
'        lblHrsDay.FontBold = True
'        lblPayID.FontBold = True
'        lblPayrollCategory.Visible = True
'        clpPayrollCategory.Visible = True
'    End If
'    If glbInsync Then
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'    End If
'    'Burlington Tech
'    If glbCompSerial = "S/N - 2351W" Then
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'        panControls.Height = 540
'    End If
'
'    'Granite Club
'    If glbCompSerial = "S/N - 2241W" Then
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'        lblTitle(9).FontBold = True
'        lblTitle(2).FontBold = True
'        lblTitle(0).FontBold = True
'        lblUnion.FontBold = True
'        lblPT.FontBold = True
'        'lblSection.FontBold = True
'        'lblReason.FontBold = True
'        'lblPayID.FontBold = True
'    End If
'
'    'Hemu - CollectCorp Inc. - Ticket #14247
'    If glbCompSerial = "S/N - 2390W" Then
'        lblHrsDay.FontBold = True
'        lblHrsWeek.FontBold = True
'        lblHrsPayPeriod.FontBold = True
'    End If
'
'    'City of Pickering - Ticket #13281
'    'The Walter Fedy Partnership - Ticket #14003
'    If glbCompSerial = "S/N - 2217W" Or glbCompSerial = "S/N - 2386W" Then
'        Dim rsTB As New ADODB.Recordset
'        rsTB.Open "SELECT ED_PT,ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
'        If Not rsTB.EOF Then
'            If rsTB("ED_PT") = "FT" Then
'                lblHrsDay.FontBold = True
'                lblHrsWeek.FontBold = True
'                lblHrsPayPeriod.FontBold = True
'            Else
'                lblHrsDay.FontBold = False
'                lblHrsWeek.FontBold = False
'                lblHrsPayPeriod.FontBold = False
'            End If
'        End If
'        rsTB.Close
'    End If
'
'    'Hamilton CAS - Ticket #13398
'    If glbCompSerial = "S/N - 2257W" Then
'        chkActPosition.Caption = "Red Circled"
'    End If
'
'    If glbLambton Then
'        lblLambtonJob.Visible = True
'        txtLambtonJob.Visible = True
'        chkUseForBenefit.DataField = "TW_USRCHECK"
'    End If
'    If glbAdv Then
'        If Not glbCompSerial = "S/N - 2242W" And Not glbCompSerial = "S/N - 2390W" Then   'london ccac
'            If isATIncluded(glbLEE_ID) Then
'                lblShift.FontBold = True
'            End If
'        End If
'        If glbLambton Then
'            txtShift.MaxLength = 4
'        End If
'    End If
'    If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
'        medBillingRate.DataField = "TW_BILLINGRATE"
'        medBillingRate.Visible = True
'        lblBillingRate.Visible = True
'    End If
'
'    If glbCompSerial = "S/N - 2380W" Then 'Vitalaire
'        lblShift.Caption = "Job Class"
'    End If
    'Ticket #16189-------------------------------
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        panControls.Height = 540
        cmdJobFiles.Visible = True
        
        If Not gSec_Inq_Job_Files_Attachment Then
            cmdJobFiles.Enabled = False
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
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
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
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
    End If
    
    If EERetrieve() = False Then
        MsgBox "Sorry, Employee can not be found"
        If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
    Else
        If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    End If
    
    'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
'        lblHrsPayPeriod.FontBold = True
'    End If
    'Ticket #16189-------------------------------
    
    Screen.MousePointer = HOURGLASS
    
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        frmETmpCrsTrnPos.Caption = IIf(glbSetPos, "Set ", "") & "Temporary/Cross Training Position History - " & Left$(glbLEE_SName, 5)
        frmETmpCrsTrnPos.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    Else
        frmETmpCrsTrnPos.Caption = "Temporary/Cross Training Position History - New Employee"
        frmETmpCrsTrnPos.lblEEName = " "
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
    Call setCaption(lblTitle(2))
    Call setCaption(lblTitle(3))
    Call setCaption(lblTitle(9))
    Call setCaption(lblSection)
    clpGrid.TABLTitle = lStr(lblGrid)
    Call Display_Value
    
    'Ticket #16189-------------------------------
'    If glbCompSerial = "S/N - 2375W" Then 'Timmins
'        rsTA.Open "SELECT ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
'        If rsTA.EOF = False And rsTA.BOF = False Then
'            If rsTA("ED_REGION") = "S" Then
'                lblHrsWeek.FontBold = True
'                lblHrsPayPeriod.FontBold = True
'            Else
'                lblHrsWeek.FontBold = False
'                lblHrsPayPeriod.FontBold = False
'            End If
'        Else
'            lblHrsWeek.FontBold = False
'            lblHrsPayPeriod.FontBold = False
'        End If
'    End If
    'Ticket #16189-------------------------------
    
    Call InitData
    Action = "M"
    savWHRS = medHours(1)
    savGrid = clpGrid.Text
    
    'Ticket #16189-------------------------------
'    If glbOttawaCCAC Then
'        frmMulti.Visible = True
'
'        Call ComEType
'        lblSection = "Emp. Type"
'        comEmpType.Visible = True
'        clpCode(5).Visible = False 'section
'        lblEndDATE.Visible = False 'end date
'        dlpENDDATE.Visible = False 'end date
'        lblReason.Visible = False 'end reason
'        clpCode(2).Visible = False 'end reason
'        frmOCCAC.Visible = True
'        frmMulti.Height = 2300
'    End If
    'Ticket #16189-------------------------------
End Sub

Private Sub Job_Desc()
Dim SQLQ As String
Dim x%
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
        'For X% = 1 To 11
        'For X% = 1 To 15
        For x% = 1 To 20
            If Not IsNull(rsWRK("JB_S" & x%)) Then JobSnap_PayScale(x) = Round2DEC(rsWRK("JB_S" & x%))
        Next
        If Not IsNull(rsWRK("JB_SALCD")) Then JobSnap_Salary_Code$ = rsWRK("JB_SALCD")
        If Not IsNull(rsWRK("JB_MIDPOINT")) Then JobSnap_MidPoint! = rsWRK("JB_MIDPOINT")
        If Not IsNull(rsWRK("JB_ORG")) Then
            clpCode(6).Visible = (rsWRK("JB_ORG") = "NONE" Or rsWRK("JB_ORG") = "EXEC") And glbWFC
        End If
    End If
Else
    clpCode(6).Visible = False
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
    Set frmETmpCrsTrnPos = Nothing
    'Ticket #16189-------------------------------
    'Call NextForm
    'Ticket #16189-------------------------------
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
    Dim SQLQ As String, Msg$, x
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
    If IsNull(dyn_HRJOBHIS("TW_ENDDATE")) Then    'New Tracking method
        dyn_HRJOBHIS("TW_CURRENT") = True
        dyn_HRJOBHIS.Update
    End If
    
    dyn_HRJOBHIS.MoveNext
    
    While Not dyn_HRJOBHIS.EOF
        If dyn_HRJOBHIS("TW_CURRENT") <> 0 Then
            dyn_HRJOBHIS("TW_CURRENT") = False
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SetCurrentFlag", "HR_TEMP_WORK", "Edit")
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
        'Ticket #16189-------------------------------
'        If fglbNew And glbCompSerial <> "S/N - 2378W" Then
'            txtPayrollID.Enabled = True
'        Else
        'Ticket #16189-------------------------------
            txtPayrollID.Enabled = False
'        End If
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
    clpJob.Enabled = TF
    txtReptAuthority(0).Enabled = TF
    txtReptAuthority(1).Enabled = TF
    txtReptAuthority(2).Enabled = TF
    
    ' sam add
    elpReptAuthShow(0).Enabled = TF
    elpReptAuthShow(1).Enabled = TF
    elpReptAuthShow(2).Enabled = TF
    txtPosCtr.Enabled = TF
    txtShift.Enabled = TF
    dlpStartDate.Enabled = TF
    txtComment.Enabled = TF     'Jaddy 6/4/99
    txtComments2.Enabled = TF
    clpCode(6).Enabled = False
    clpPayrollCategory.Enabled = TF
    dlpENDDATE.Enabled = TF
    clpCode(2).Enabled = TF
    
    'If Not gSec_Inq_Salary Then cmdSalary.Enabled = False
    'If Not gSec_Inq_Performance Then cmdPerform.Enabled = False
    
    'George on Jan 26,2006 #10266
    glbJob = "" 'George on Jan 24,2006 #10266
    glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If Not IsNull(Data1.Recordset("TW_JOB")) Then glbJob = Data1.Recordset("TW_JOB") 'George on Jan 19,2006 #10266
        If Not IsNull(Data1.Recordset("TW_SDATE")) Then glbSDate = Data1.Recordset("TW_SDATE") 'George on Jan 24,2006 #10266
    End If

    'Ticket #16189-------------------------------
'    glbDocName = "Offer"
'    If gsAttachment_DB Then
'        Call DispimgIcon(Me, "frmETmpCrsTrnPos")
'        If gSec_Upd_Position And Not glbtermopen Then
'            If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'                cmdImport.Visible = False
'            Else
'                cmdImport.Visible = True
'            End If
'        End If
'    End If
'    'George on Jan 26,2006 #10266
'
'    'Simona - begin - Assessment Strategies-#14963
'    If (glbCompSerial = "S/N - 2401W") Then
'        If NewHireForms.count > 0 Then
'            medHours(0).Text = "7.5"
'            medHours(1).Text = "37.5"
'            medHours(2).Text = "75.0"
'            medFTENum.Text = "1"
'            medFTEHrs.Text = "1950"
'        End If
'    End If
'    'Simona - end - Assessment Strategies-#14963
'
'    If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #15590
'        lblTitle(3).Enabled = False
'        clpGLNum.Enabled = False
'    End If
    'Ticket #16189-------------------------------
    
End Sub

Private Sub medHours_LostFocus(Index As Integer)
    'Ticket #16189-------------------------------
'    'City of Pickering - Ticket #13281
'    If glbCompSerial = "S/N - 2217W" Then
'        If lblHrsDay.FontBold = True Or lblHrsWeek.FontBold = True Or lblHrsPayPeriod.FontBold = True Then
'            If IsNumeric(medHours(2)) Then  'Hours/Pay Period
'                medFTEHrs = medHours(2) * 26
'            End If
'        End If
'    End If
    'Ticket #16189-------------------------------
End Sub

Private Sub scrControl_Change()
    fraDetail.Top = 240 + vbxTrueGrid.Height + panEEDESC.Height - scrControl.Value * ((panControls.Height + scrControl.Max) / scrControl.Max)
End Sub

Private Sub txtComment_GotFocus()
    Call SetPanHelp(Me.ActiveControl)       'Jaddy 6/4/99
End Sub

Private Sub txtComments2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
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
End Sub

Private Sub txtShift_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

'Ticket #16189-------------------------------
'Private Sub Upd_Related_Salary()
'
'Dim SQLQ As String, Msg As String
'Dim dynHRSALHIS As New ADODB.Recordset
'Dim JobCode$, PositionStartDat, JobReason$
'Dim HoursPerWeek!
'Dim lngJobID&
'
'Dim x!, cX$
'Dim SH_SALARY@, SH_SALCD$, SH_EDATE, SH_PAYP$, SH_NEXTDAT As Variant
'Dim xSH_FISCALYEAR, xSH_SECTION, xSH_MARKETLINE, xSH_BAND 'WFC ONLY
'Dim SHisDate, SPosDate  As Variant
'Dim AnnualSalary As Double, Compa!, SalaryGrade$
'Dim xPosEarly
'Dim xSH_PREMIUM, xSH_TOTAL, xSH_VGROUP, xSH_VSTEP
'Dim xSHID 'George added Mar 9,2006 #9965
'On Error GoTo UpRel_Err
'
'JobCode$ = clpJob.Text
'
'If IsNumeric(Data1.Recordset("TW_ID")) Then lngJobID& = Data1.Recordset("TW_ID") Else lngJobID& = 0
'
'If Not IsNull(dlpStartDate.Text) Then PositionStartDat = CVDate(dlpStartDate.Text)
'If Not IsNull(medHours(1)) And Len(medHours(1)) > 0 Then HoursPerWeek! = medHours(1)
'If Not IsNull(clpCode(1).Text) Then JobReason$ = clpCode(1).Text
'
'
'SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
'SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
'SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL Or glbOracle, "DESC", "")
'dynHRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'If dynHRSALHIS.BOF And dynHRSALHIS.EOF Then
'    Msg = "No salary records found - New Employee?" & Chr(10)
'    Msg = Msg & "Please review and update this Employee's" & Chr(10)
'    Msg = Msg & "salary."
'    MsgBox Msg
'    dynHRSALHIS.Close
'    Exit Sub
'End If
'
'SHisDate = CVDate(dynHRSALHIS("SH_EDATE"))
'SPosDate = CVDate(dynHRSALHIS("SH_SDATE"))
'xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0
'If xPosEarly Then
'    If fgtxtStartDate = SHisDate And dynHRSALHIS("SH_JOB") = JobCode$ Then
'        dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
'        dynHRSALHIS.Update
'        Exit Sub
'    End If
'End If
'
'dynHRSALHIS("SH_CURRENT") = False
'dynHRSALHIS.Update
'xSHID = dynHRSALHIS("SH_ID")
''George added Mar 9,2006 #9965
''If glbCompSerial = "S/N - 2259W" Or glbGP Then
''    Call Salary_Integration(glbLEE_ID, , False, False, xSHID)
''End If
''George added Mar 9,2006 #9965
'
'If Not IsNull(dynHRSALHIS.Fields("SH_SALARY")) Then SH_SALARY@ = dynHRSALHIS.Fields("SH_SALARY")
'If Not IsNull(dynHRSALHIS.Fields("SH_SALCD")) Then SH_SALCD$ = dynHRSALHIS.Fields("SH_SALCD")
'If Not IsNull(dynHRSALHIS.Fields("SH_PAYP")) Then SH_PAYP$ = dynHRSALHIS.Fields("SH_PAYP")
'If Not IsNull(dynHRSALHIS.Fields("SH_NEXTDAT")) Then SH_NEXTDAT = dynHRSALHIS.Fields("SH_NEXTDAT")
'If glbWFC Then
'    xSH_FISCALYEAR = "": xSH_SECTION = "": xSH_MARKETLINE = "": xSH_BAND = ""
'    If Not IsNull(dynHRSALHIS.Fields("SH_FISCALYEAR")) Then xSH_FISCALYEAR = dynHRSALHIS.Fields("SH_FISCALYEAR")
'    If Not IsNull(dynHRSALHIS.Fields("SH_SECTION")) Then xSH_SECTION = dynHRSALHIS.Fields("SH_SECTION")
'    If Not IsNull(dynHRSALHIS.Fields("SH_MARKETLINE")) Then xSH_MARKETLINE = dynHRSALHIS.Fields("SH_MARKETLINE")
'    If Len(clpCode(6).Text) > 0 Then xSH_BAND = clpCode(6).Text
'End If
'If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
'    If Not IsNull(dynHRSALHIS.Fields("SH_PREMIUM")) Then xSH_PREMIUM = dynHRSALHIS.Fields("SH_PREMIUM")
'    If Not IsNull(dynHRSALHIS.Fields("SH_TOTAL")) Then xSH_TOTAL = dynHRSALHIS.Fields("SH_TOTAL")
'    If Not IsNull(dynHRSALHIS.Fields("SH_VGROUP")) Then xSH_VGROUP = dynHRSALHIS.Fields("SH_VGROUP")
'    If Not IsNull(dynHRSALHIS.Fields("SH_VSTEP")) Then xSH_VSTEP = dynHRSALHIS.Fields("SH_VSTEP")
'End If
'
''SET COMPA RATIO
''================
''Days and Months added by Bryan 30/Sep/05 Ticket#9354
'If JobSnap_Salary_Code$ = "A" Then
'    If SH_SALCD$ = "H" Then
'        AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52
'    ElseIf SH_SALCD$ = "M" Then
'        AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 12
'    ElseIf SH_SALCD$ = "D" Then
'        If GetLeapYear(Year(Date)) Then
'            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 366
'        Else
'            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 265
'        End If
'    Else
'        AnnualSalary = SH_SALARY@
'    End If
'ElseIf JobSnap_Salary_Code$ = "H" Then
'    If SH_SALCD$ = "A" Then
'        If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 52
'    ElseIf SH_SALCD$ = "M" Then
'        If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
'    ElseIf SH_SALCD$ = "D" Then
'        If GetLeapYear(Year(Date)) Then
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 366
'        Else
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 365
'        End If
'    Else
'        AnnualSalary = SH_SALARY@
'    End If
'ElseIf JobSnap_Salary_Code$ = "M" Then
'    If SH_SALCD$ = "A" Then
'        AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
'    ElseIf SH_SALCD$ = "M" Then
'        AnnualSalary = SH_SALARY@
'    ElseIf SH_SALCD$ = "D" Then
'        If GetLeapYear(Year(Date)) Then
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 366) / 12
'        Else
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 365) / 12
'        End If
'    Else
'        If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52 / 12
'    End If
'ElseIf JobSnap_Salary_Code$ = "D" Then
'    If SH_SALCD$ = "H" Then
'        If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) / 52
'    ElseIf SH_SALCD$ = "M" Then
'        If GetLeapYear(Year(Date)) Then
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 366
'        Else
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 365
'        End If
'    ElseIf SH_SALCD$ = "A" Then
'        If GetLeapYear(Year(Date)) Then
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 366
'        Else
'            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 365
'        End If
'    Else
'        AnnualSalary = SH_SALARY@
'    End If
'End If
' ' set COMPA RATIO
'If JobSnap_PayScale(JobSnap_MidPoint!) <> 0 And AnnualSalary <> 0 Then
'    Compa! = (AnnualSalary / JobSnap_PayScale(JobSnap_MidPoint!)) * 100
'Else
'    Compa! = 0
'End If
'
'If Compa! > 999.99 Then
'    Compa! = 999.99
'End If
''Determine Pay Scale individual fits into
''==========================================
'SalaryGrade$ = "00"
'For x! = 1 To 11
'    If AnnualSalary >= JobSnap_PayScale(x) And JobSnap_PayScale(x) > 0 Then
'      cX$ = CStr(x)
'      If x! <= 9 Then cX$ = "0" & cX$
'      SalaryGrade$ = cX$
'    End If
'Next x!
''NOW UPDATE SALARY HISTORY TABLE  - only if new record do we add record
''================================
'If DateDiff("d", PositionStartDat, SHisDate) > 0 And glbSetPos Then GoTo SkipSal_Change
'
'If Not xPosEarly Then dynHRSALHIS.AddNew
'
'dynHRSALHIS("SH_COMPNO") = "001" 'SH_COMPNO%
'dynHRSALHIS("SH_EMPNBR") = glbLEE_ID
'dynHRSALHIS("SH_CURRENT") = True
'dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
'dynHRSALHIS("SH_EDATE") = IIf(xPosEarly, SHisDate, CVDate(PositionStartDat))
'dynHRSALHIS("SH_TRANSDATE") = Format(Now, "SHORT DATE")
'dynHRSALHIS("SH_SALARY") = SH_SALARY@
'dynHRSALHIS("SH_SALCD") = SH_SALCD$
'dynHRSALHIS("SH_JOB") = JobCode$
'dynHRSALHIS("SH_GRID") = clpGrid.Text
'dynHRSALHIS("SH_PAYROLL_ID") = txtPayrollID
''lngJobID&
'dynHRSALHIS("SH_JOB_ID") = lngJobID&
'dynHRSALHIS("SH_PAYP_TABLE") = "SDPP"
'dynHRSALHIS("SH_PAYP") = SH_PAYP$
'If IsDate(SH_NEXTDAT) Then
'    If CVDate(SH_NEXTDAT) > IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)) Then
'        dynHRSALHIS("SH_NEXTDAT") = SH_NEXTDAT
'    End If
'End If
'dynHRSALHIS("SH_WHRS") = HoursPerWeek!
'dynHRSALHIS("SH_SREAS_TABLE") = "SDRC"
'dynHRSALHIS("SH_SREAS1") = JobReason$     ' reason code
'dynHRSALHIS("SH_COMPA") = Round(Compa!, 2)
'dynHRSALHIS("SH_GRADE") = SalaryGrade$
'dynHRSALHIS("SH_LDATE") = Date
'dynHRSALHIS("SH_LTIME") = Time$
'dynHRSALHIS("SH_LUSER") = glbUserID
'If glbWFC Then
'    If Len(xSH_FISCALYEAR) > 0 Then
'        dynHRSALHIS("SH_FISCALYEAR") = xSH_FISCALYEAR
'    End If
'    If Len(xSH_SECTION) > 0 Then
'        dynHRSALHIS("SH_SECTION") = xSH_SECTION
'    End If
'    If Len(xSH_MARKETLINE) > 0 Then
'        dynHRSALHIS("SH_MARKETLINE") = xSH_MARKETLINE
'    End If
'    If Len(xSH_BAND) > 0 Then
'        dynHRSALHIS("SH_BAND") = xSH_BAND
'    End If
'End If
'If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
'    If Len(xSH_PREMIUM) > 0 Then
'        dynHRSALHIS("SH_PREMIUM") = xSH_PREMIUM
'    End If
'    If Len(xSH_TOTAL) > 0 Then
'        dynHRSALHIS("SH_TOTAL") = xSH_TOTAL
'    End If
'    If Len(xSH_VGROUP) > 0 Then
'        dynHRSALHIS("SH_VGROUP") = xSH_VGROUP
'    End If
'    If Len(xSH_VSTEP) > 0 Then
'        dynHRSALHIS("SH_VSTEP") = xSH_VSTEP
'    End If
'End If
'dynHRSALHIS.Update
'SkipSal_Change:
'xSHID = dynHRSALHIS("SH_ID")
'dynHRSALHIS.Close
'Call updBenefitForSalDEPN(glbLEE_ID)
'
''City of Niagara Falls - Ticket #15542
'If glbVadim And glbCompSerial = "S/N - 2276W" Then
'    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
'    Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)), "", Val(SalaryGrade$), JobCode$, "A")
'End If
'
''George added Mar 9,2006 #9965
'If glbCompSerial = "S/N - 2259W" Or glbGP Then 'Or (glbWFC And glbPlantCode = "GREN") Then
'    Call Salary_Integration(glbLEE_ID, , False, IIf(xPosEarly, False, True), xSHID)
'End If
''George added Mar 9,2006 #9965
'
'Exit Sub
'
'UpRel_Err:
'If Err = 3021 Then
'    Exit Sub
'End If
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HRSAL/PERF", "INSERT")
'Call RollBack '26July99 js
'
'End Sub
''Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
''If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
''    KeyAscii = 0
''    Exit Sub
''End If
''If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
''End Sub
'Ticket #16189-------------------------------

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
    Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim x As Integer
    
    Call Display_Value
    Call Job_Desc
    
    'George on Dec 7,2005 #9928 begin
    If chkCurrent(0) Then
        chkActPosition.Enabled = True
    Else
        chkActPosition.Enabled = False
    End If
    'George on Dec 7,2005 #9928 end

    'glbJob = Data1.Recordset("TW_JOB") 'George on Jan 19,2006 #10266
    'glbSDate = Data1.Recordset("TW_SDATE") 'George on Jan 24,2006 #10266

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
    Dim strNUM As String, x%

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

'Ticket #16189-------------------------------
'Private Function updFollow(xType)
'Dim newline As String
'Dim SQLQ As String
'Dim Msg As String
'Dim rsTB As New ADODB.Recordset
'Dim dynHRAT As New ADODB.Recordset
'Dim Edit1 As Integer
'
'newline = Chr$(13) & Chr$(10)
'updFollow = False
'
'On Error GoTo CrFollow_Err
'
'If oENDDATE <> "" Then    'DATE Renewal IS NOW MANDATORY
'    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND EF_FREAS = 'RFED'"
'    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(oENDDATE)
'    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If dynHRAT.BOF And dynHRAT.EOF Then
'        Edit1 = False
'    Else
'        Edit1 = True    ' returns true if found records
'    End If
'Else
'    Edit1 = False
'End If
'
'If xType = "U" Then
'
'    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'    If fglbNew And dlpENDDATE.Text <> "" Then
'        rsTB.AddNew
'        rsTB("EF_COMPNO") = "001"
'        rsTB("EF_EMPNBR") = glbLEE_ID
'        rsTB("EF_FDATE") = CVDate(dlpENDDATE.Text)
'        rsTB("EF_FREAS_TABL") = "FURE"
'        rsTB("EF_ADMINBY_TABL") = "EDAB"
'        rsTB("EF_FREAS") = "RFED"
'        rsTB("EF_COMMENTS") = ""
'        rsTB("EF_LDATE") = Date
'        rsTB("EF_LTIME") = Time$
'        rsTB("EF_LUSER") = glbUserID
'        rsTB.Update
'        rsTB.Close
'        updFollow = True
'        Msg = "A Follow Up Record was created!"
'        'MsgBox Msg
'        Exit Function
'    End If
'    If fglbNew% = False And Edit1 = False And dlpENDDATE.Text <> "" Then
'        rsTB.AddNew
'        rsTB("EF_COMPNO") = "001"
'        rsTB("EF_EMPNBR") = glbLEE_ID
'        rsTB("EF_FDATE") = CVDate(dlpENDDATE.Text)
'        rsTB("EF_FREAS_TABL") = "FURE"
'        rsTB("EF_ADMINBY_TABL") = "EDAB"
'        rsTB("EF_FREAS") = "RFED"
'        rsTB("EF_COMMENTS") = ""
'        rsTB("EF_LDATE") = Date
'        rsTB("EF_LTIME") = Time$
'        rsTB("EF_LUSER") = glbUserID
'        rsTB.Update
'        rsTB.Close
'        updFollow = True
'        Msg = "A Follow Up Record was created!"
'        'MsgBox Msg
'        Exit Function
'    End If
'
'    If fglbNew% = False And Edit1 = True And dlpENDDATE.Text <> "" Then ' edited record
'        'EOF?
'        dynHRAT.MoveFirst
'        Do Until dynHRAT.EOF
'            'dynHRAT.Edit
'            dynHRAT("EF_COMPNO") = "001"
'            dynHRAT("EF_EMPNBR") = glbLEE_ID
'            dynHRAT("EF_FDATE") = CVDate(dlpENDDATE.Text)
'            dynHRAT("EF_FREAS") = "RFED"
'            dynHRAT("EF_COMMENTS") = ""
'            dynHRAT("EF_LDATE") = Date
'            dynHRAT("EF_LTIME") = Time$
'            dynHRAT("EF_LUSER") = glbUserID
'            dynHRAT.Update
'            dynHRAT.MoveNext
'        Loop
'        dynHRAT.Close
'        If oENDDATE <> dlpENDDATE.Text Then
'            Msg = "A Follow Up Record was updated!"
'            'MsgBox Msg
'        End If
'        updFollow = True
'        Edit1 = True
'        Exit Function
'    End If
'    If fglbNew% = False And Edit1 = True And dlpENDDATE.Text = "" Then
'        Do Until dynHRAT.EOF
'            dynHRAT.Delete
'            dynHRAT.MoveNext
'        Loop
'        dynHRAT.Close
'        Edit1 = True
'        updFollow = True
'        Msg = "A record has been deleted from the Follow Up table"
'        'MsgBox Msg
'        Exit Function
'    End If
'Else
'    If Edit1 = True Then
'        Do Until dynHRAT.EOF
'            dynHRAT.Delete
'            dynHRAT.MoveNext
'        Loop
'        dynHRAT.Close
'        Edit1 = True
'        updFollow = True
'        Msg = "A record has been deleted from the Follow Up table"
'        'MsgBox Msg
'        Exit Function
'    Else
'        updFollow = True
'    End If
'End If
'
'If dlpENDDATE.Text = "" Then
'    updFollow = True
'End If
'
'Exit Function
'
'CrFollow_Err:
'If Err = 3022 Then
'    MsgBox "The record is not entered or deleted!"
'    Err = 0   ' i know will be reset any way - but just in case
'    Resume Next
'    Exit Function
'End If
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
'Resume Next
'
'End Function

'Function EMPBenefitGroup()
'    Dim rsEmp As New ADODB.Recordset
'    Dim SQLQ
'
'    EMPBenefitGroup = ""
'    SQLQ = "Select ED_BENEFIT_GROUP from HREMP Where ED_EMPNBR = " & glbLEE_ID
'    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'    If Not rsEmp.EOF Then
'        EMPBenefitGroup = rsEmp("ED_BENEFIT_GROUP") & ""
'    End If
'End Function

'Sub AddFTE(xEmpNo, xFLAG)
'    Dim OldFTE, NewFTE, xEFDATE, xETDATE, xNumVac
'    Dim RsFTEHis As New ADODB.Recordset
'    Dim xDays1, xDays2, xVacDays, xDate1, xDate2, xFDate, xTDate, xHrsDay, xHrsDayN
'    Dim xVacHours, xYear, xNum As Integer, II, J
'    Dim xArray(100, 2)
'    Dim tNewFTE, xNumVacINS, VAC_First
'    Dim RsTempEmp As New ADODB.Recordset
'    Dim RsJobEmp As New ADODB.Recordset
'    Dim SQLQ, xTxtJOB
'    Dim FlagLoop As Boolean
'
'    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
'    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    xEFDATE = ""
'    xETDATE = ""
'    xNumVac = 0
'    If Not RsTempEmp.EOF Then
'        xNumVac = RsTempEmp("ED_VAC")
'        xNumVacINS = RsTempEmp("ED_VAC")
'        xEFDATE = RsTempEmp("ED_EFDATE")
'        xETDATE = RsTempEmp("ED_ETDATE")
'    End If
'    RsTempEmp.Close
'
'    If Len(xEFDATE) = 0 Or Len(xETDATE) = 0 Then
'        Exit Sub
'    End If
'
'    'If xFLAG = "NEW" Then
'        Call Pause(6)
'    'End If
'    SQLQ = "Select * from HR_JOB_HISTORY Where TW_EMPNBR = " & xEmpNo
'    SQLQ = SQLQ & " ORDER BY TW_SDATE DESC"
'    RsJobEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If RsJobEmp.EOF Then
'        Exit Sub
'    End If
'
'    SQLQ = "SELECT * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
'    If IsDate(xEFDATE) Then
'    SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
'    End If
'    If IsDate(xETDATE) Then
'    SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
'    End If
'    SQLQ = SQLQ & "ORDER BY CP_FDATE DESC"
'    RsFTEHis.Open SQLQ, gdbAdoSN2322, adOpenKeyset, adLockOptimistic
'    If RsFTEHis.EOF And xFLAG <> "NEW" Then
'        Exit Sub
'    End If
'
'    If xFLAG = "NEW" Then
'        If xNumVac = 0 Then
'            Exit Sub
'        End If
'        If Not RsFTEHis.EOF Then ' IF CP_VACORIGION EXIST AND CHANGE IN THE SAME YEAR
'            If RsFTEHis("CP_FDATE") = xEFDATE Then
'                xNumVac = RsFTEHis("CP_VACORIGION")
'                GoTo MAIN_DEAL
'            End If
'        End If
'        '' The following shows how to calculate the VAC days at the end of last year
'        '' We always suppose the FTE# is 1.00 at the end of last year
'        ' X is VAC days when FTE# = 1
'        ' VAC_First is the first VAC days before FTE# change
'        ' days1,days2, ... daysn are date range when FTE# change within this year
'        ' VAC_First = X/365 * FTE#1 * days1 + X/365 * FTE#2 * days2 + ... + X/365 * FTE#n * daysn
'        ' X = (VAC_First * 365)/(FTE#1 * days1 + FTE#2 * days2 + ... + FTE#n * daysn)
'        VAC_First = xNumVac
'
'        xDate1 = "**"
'        xFDate = xEFDATE
'        xTDate = xETDATE
'        FlagLoop = True
'        xHrsDayN = 0
'        If RsJobEmp("TW_DHRS") = 0 Then
'            xHrsDayN = 0
'        Else
'            If IsNull(RsJobEmp("TW_DHRS")) Then
'                xHrsDayN = 0
'            Else
'                xHrsDayN = RsJobEmp("TW_DHRS")
'            End If
'        End If
'
'        RsJobEmp.MoveNext
'        II = 0
'        Do While (Not RsJobEmp.EOF) And FlagLoop
'            xDate1 = RsJobEmp("TW_SDATE")
'            If CVDate(xDate1) > CVDate(xETDATE) Then
'                GoTo Next_Rec00
'            End If
'            If RsJobEmp("TW_FTENUM") = 0 Then
'                GoTo Next_Rec00
'            End If
'            If IsNull(RsJobEmp("TW_FTENUM")) Then
'                GoTo Next_Rec00
'            End If
'            OldFTE = RsJobEmp("TW_FTENUM")
'
'            If RsJobEmp("TW_DHRS") = 0 Then
'                GoTo Next_Rec00
'            End If
'            If IsNull(RsJobEmp("TW_DHRS")) Then
'                GoTo Next_Rec00
'            End If
'            xHrsDay = RsJobEmp("TW_DHRS")
'
'            If CVDate(xDate1) < CVDate(xEFDATE) Then
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate)) * OldFTE
'                FlagLoop = False
'            Else
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
'                xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
'            End If
'
'Next_Rec00:
'            RsJobEmp.MoveNext
'        Loop
'        If IsDate(xDate1) Then
'            If CVDate(xDate1) > CVDate(xEFDATE) Then
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
'            End If
'        End If
'
'        xVacDays = 0
'        For J = 1 To II
'            xVacDays = xVacDays + xArray(J, 1)
'        Next
'        If xVacDays = 0 Then
'            Exit Sub
'        End If
'        If xHrsDay = 0 Then
'            Exit Sub
'        End If
'        xNumVac = Round((((VAC_First * 365) / (xVacDays)) / xHrsDayN), 0) * xHrsDayN
''        If RsFTEHis.EOF Then
''            RsFTEHis.AddNew
''        End If
''        RsFTEHis("CP_EMPNBR") = xEmpNo
''        RsFTEHis("CP_VACORIGION") = xNumVac
''        RsFTEHis("CP_VACO") = xNumVacINS
''        RsFTEHis("CP_FDATE") = xEFDATE
''        RsFTEHis("CP_TDATE") = xETDATE
''        RsFTEHis("CP_LDATE") = DATE
''        RsFTEHis("CP_LTIME") = Time$
''        RsFTEHis("CP_LUSER") = glbUSERID
''        RsFTEHis.Update
'    End If
'
'    If xFLAG <> "NEW" Then
'        If RsFTEHis.EOF Then
'            xNumVac = 0
'            Exit Sub
'        Else
'            xNumVac = RsFTEHis("CP_VACORIGION")
'        End If
'    End If
'
'    '--- Above Got vacation days per year when FTE = 1 (xNumVac)
'MAIN_DEAL:
'    II = 0
'    xDate1 = "**"
'    xFDate = xEFDATE
'    xTDate = xETDATE
'    FlagLoop = True
'    RsJobEmp.MoveFirst
'    Do While (Not RsJobEmp.EOF) And FlagLoop
'        xDate1 = RsJobEmp("TW_SDATE")
'        If CVDate(xDate1) > CVDate(xETDATE) Then
'            GoTo Next_Rec01
'        End If
'        If RsJobEmp("TW_FTENUM") = 0 Then
'            GoTo Next_Rec01
'        End If
'        If IsNull(RsJobEmp("TW_FTENUM")) Then
'            GoTo Next_Rec01
'        End If
'        OldFTE = RsJobEmp("TW_FTENUM")
'
'        If RsJobEmp("TW_DHRS") = 0 Then
'            GoTo Next_Rec01
'        End If
'        If IsNull(RsJobEmp("TW_DHRS")) Then
'            GoTo Next_Rec01
'        End If
'        xHrsDay = RsJobEmp("TW_DHRS")
'
'        If CVDate(xDate1) < CVDate(xEFDATE) Then
'            II = II + 1
'            xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate))
'            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
'            FlagLoop = False
'        Else
'            II = II + 1
'            xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
'            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
'            xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
'
'        End If
'
'Next_Rec01:
'        RsJobEmp.MoveNext
'    Loop
'    'If IsDate(xDate1) Then
'    '    If CVDate(xDate1) > CVDate(xEFDATE) Then
'    '        II = II + 1
'    '        xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
'    '        xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
'    '    End If
'    'End If
'
'    xVacDays = 0
'    For J = 1 To II
'        xVacDays = xVacDays + xArray(J, 2)
'    Next
'
'    If xVacDays = 0 Then
'        Exit Sub
'    End If
'    'xVacHours = Round(xVacDays, 0) * xHrsDay
'    xVacHours = Round25(xVacDays) * xHrsDay
'
'    Call Pause(0.5) 'Add By Frank August 1, 2001
'
'    'Dim RsTempEmp As New ADODB.Recordset
'    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
'    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'    If Not RsTempEmp.EOF Then
'        RsTempEmp("ED_VAC") = xVacHours
'        RsTempEmp.Update
'
'        If xFLAG = "NEW" Then
'            'If RsFTEHis.EOF Then
'            RsFTEHis.AddNew
'            RsFTEHis("CP_EMPNBR") = xEmpNo
'            RsFTEHis("CP_VACORIGION") = xNumVac
'            RsFTEHis("CP_VACO") = xNumVacINS
'            RsFTEHis("CP_VACN") = xVacHours
'            RsFTEHis("CP_FTENUMO") = fOldFTE
'            RsFTEHis("CP_FTENUMN") = fNewFTE
'            RsFTEHis("CP_FDATE") = CVDate(xEFDATE)
'            RsFTEHis("CP_TDATE") = CVDate(xETDATE)
'            RsFTEHis("CP_LDATE") = Date
'            RsFTEHis("CP_LTIME") = Time$
'            RsFTEHis("CP_LUSER") = glbUserID
'            RsFTEHis.Update
'            RsFTEHis.Close
'            'End If
'        Else
'            RsFTEHis.Close
'            If fOldFTE > 0 Then
'            SQLQ = "DELETE * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
'            SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
'            SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
'            SQLQ = SQLQ & "AND CP_VACN = " & xNumVacINS & " "
'            SQLQ = SQLQ & "AND CP_FTENUMN = " & fOldFTE & " "
'            gdbAdoSN2322.Execute SQLQ
'            End If
'        End If
'    End If
'    RsTempEmp.Close
'
'
'    Exit Sub
'
'ExitLin1:
'End Sub

'Private Function Round25(xNumb)
'Dim xInteger, xDecimal, xDecTmp
'    xInteger = Int(xNumb)
'    xDecimal = xNumb - xInteger
'    xDecTmp = 0
'    If xDecimal >= 0 And xDecimal < 0.25 Then
'        xDecTmp = 0
'    End If
'    If xDecimal >= 0.25 And xDecimal < 0.75 Then
'        xDecTmp = 0.5
'    End If
'    If xDecimal >= 0.75 Then
'        xDecTmp = 1
'    End If
'    Round25 = xInteger + xDecTmp
'End Function
'Ticket #16189-------------------------------

''' Sam add July 2002 * Remove Binding Control
Public Sub Display_Value()
    Dim SQLQ
    Dim x, xFld
    
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
    Else
        If glbtermopen Then
            SQLQ = "Select Term_TEMP_WORK.*"
        Else
            SQLQ = "Select HR_TEMP_WORK.*"
        End If
        
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_TEMP_WORK"
            SQLQ = SQLQ & " WHERE TW_ID = " & Data1.Recordset!TW_ID
            SQLQ = SQLQ & " ORDER BY "
            If glbMulti Then SQLQ = SQLQ & "TW_CURRENT " & IIf(glbSQL, "DESC", "") & ","
            SQLQ = SQLQ & "TW_SDATE DESC"
            
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            
            If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
                rsDATA.CursorLocation = adUseServer
            End If
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = SQLQ & " FROM HR_TEMP_WORK"
            SQLQ = SQLQ & " WHERE TW_ID = " & Data1.Recordset!TW_ID
            SQLQ = SQLQ & " ORDER BY "
            If glbMulti Then SQLQ = SQLQ & "TW_CURRENT " & IIf(glbSQL, "DESC", "") & ","
            SQLQ = SQLQ & "TW_SDATE DESC"
            
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
                rsDATA.CursorLocation = adUseServer 'Oracle version needs this
            End If
            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        
        If glbLinamar Then
            clpJob.TransDiv = Right(glbLEE_ID, 3)
            'If rsDATA("TW_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1
        ElseIf glbMulti Then 'George on Dec 7,2005 #9928 begin
            'If rsDATA("TW_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1 'George on Dec 7,2005 #9928 end
        End If
        
        If rsDATA("TW_POSITION_CONTROL") & "" = "YES" Then chkActPosition = 1
        
        Call Set_Control("R", Me, rsDATA)
        
        'Friesens - Ticket #16189
        If glbCompSerial = "S/N - 2279W" Then
            cmdJobFiles.Enabled = True
            
            If Not gSec_Inq_Job_Files_Attachment Then
                cmdJobFiles.Enabled = False
            End If
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
    UpdateRight = gSec_Upd_Temp_Cross_Training
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
        frmETmpCrsTrnPos.Caption = "Temporary/Cross Training Assignment - " & Left$(glbLEE_SName, 5)
        frmETmpCrsTrnPos.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

'Ticket #16189-------------------------------
'Private Function UpdPositionCCAC()
'Dim rsOC As New ADODB.Recordset
'Dim rsJOBOC As New ADODB.Recordset
'Dim SQLQ
'If glbOttawaCCAC Then
'
'    SQLQ = "SELECT TW_ID FROM HR_JOB_HISTORY WHERE TW_EMPNBR=" & glbLEE_ID
'    SQLQ = SQLQ & " AND TW_SDATE>" & Date_SQL(dlpStartDate)
'
'    rsJOBOC.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'    If Not rsJOBOC.EOF Then Exit Function
'
'    UpdPositionCCAC = False
'    SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_EMPNBR =" & glbLEE_ID
'    rsOC.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
'    If Not rsOC.EOF Then
'        rsOC("PC_EMPNBR") = Null
'        rsOC.Update
'    End If
'    rsOC.Close
'
'    If Len(txtPosCtr) > 0 Then
'        SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_JOB='" & clpJob.Text & "'"
'        SQLQ = SQLQ & " AND PC_POSITION_CONTROL='" & txtPosCtr & "'"
'        rsOC.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
'        If rsOC.EOF Then
'            rsOC.Close
'            MsgBox "Invalid CCAC Position Number!"
'            txtPosCtr.SetFocus
'            Exit Function
'        Else
'            If IsNull(rsOC("PC_EMPNBR")) Then
'                rsOC("PC_EMPNBR") = glbLEE_ID
'                rsOC.Update
'            Else
'                MsgBox "The CCAC Position number has already been used."
'                txtPosCtr.SetFocus
'                Exit Function
'            End If
'            rsOC.Close
'        End If
'    End If
'End If
'UpdPositionCCAC = True
'End Function
'Ticket #16189-------------------------------

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
    
    'Ticket #16189-------------------------------
'    If glbOttawaCCAC Then
'        rsEmp.Open "SELECT ED_DEPTNO,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP,ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
'        If Not rsEmp.EOF Then
'            clpDept = Format(rsEmp("ED_DEPTNO"), "@")
'            clpDiv = Format(rsEmp("ED_DIV"), "@")
'            clpGLNum = Format(rsEmp("ED_GLNO"), "@")
'            clpCode(4) = Format(rsEmp("ED_EMP"), "@")
'            clpCode(0) = Format(rsEmp("ED_ORG"), "@")
'            clpPT = Format(rsEmp("ED_PT"), "@")
'            If glbCompSerial = "S/N - 2332W" Then
'                clpCode(5) = Format(rsEmp("ED_SECTION"), "@")
'            Else
'                txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@")
'            End If
'        End If
'    Else
    'Ticket #16189-------------------------------
    If glbMulti Or glbVadim Then
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
        txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@")
        txtPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@")
        clpCode(5) = Format(rsEmp("ED_SECTION"), "@")
        If glbCompSerial = "S/N - 2362W" Then   'city of sarnia
            clpPayrollCategory = clpDiv
        End If
        If glbCompSerial = "S/N - 2363W" Then ' CITY OF K LAKES
            clpPayrollCategory = rsEmp("ED_REGION") & ""
        End If
        empPayrollID = txtPayrollID
        rsEmp.Close
    End If
    
    'Ticket #16189-------------------------------
'     'Simona - begin - Assessment Strategies-#14963
'    If (glbCompSerial = "S/N - 2401W") Then
'        If NewHireForms.count > 0 Then
'            medHours(0).Text = "7.5"
'            medHours(1).Text = "37.5"
'            medHours(2).Text = "75.0"
'            medFTENum.Text = "1"
'            medFTEHrs.Text = "1950"
'        End If
'    End If
'    'Simona - end - Assessment Strategies-#14963
    'Ticket #16189-------------------------------
End Sub

'Ticket #16189-------------------------------
'Private Sub SetEmpValue(Optional ReSetOldValue As Boolean)
'    Dim rsEmp As New ADODB.Recordset
'    Dim xUpdate As Boolean
'
'    rsEmp.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP,ED_PAYROLL_ID,ED_SECTION,ED_REGION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    If rsEmp.EOF Then Exit Sub
'    If ReSetOldValue Then
'        oDeptNo = Format(rsEmp("ED_DEPTNO"), "@")
''        ODIV = Format(rsEMP("ED_DIV"), "@")
'        oGLNo = Format(rsEmp("ED_GLNO"), "@")
'        oStatus = Format(rsEmp("ED_EMP"), "@")
'        oOrg = Format(rsEmp("ED_ORG"), "@")
''        oPT = Format(rsEMP("ED_PT"), "@")
''        OEmptype = Format(rsEMP("ED_EMPTYPE"), "@")
'        oPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@")
'        'OSection = Format(rsEMP("ED_SECTION"), "@")
'    End If
'
'    xUpdate = False
'    If clpDept <> Format(rsEmp("ED_DEPTNO"), "@") Then xUpdate = True
'    If clpDiv = Format(rsEmp("ED_DIV"), "@") Then xUpdate = True
'    If clpGLNum = Format(rsEmp("ED_GLNO"), "@") Then xUpdate = True
'    If clpCode(4) = Format(rsEmp("ED_EMP"), "@") Then xUpdate = True
'    If clpCode(0) = Format(rsEmp("ED_ORG"), "@") Then xUpdate = True
'    If clpPT = Format(rsEmp("ED_PT"), "@") Then xUpdate = True
'    If txtEmpType = Format(rsEmp("ED_EMPTYPE"), "@") Then xUpdate = True
'    If txtPayrollID = Format(rsEmp("ED_PAYROLL_ID"), "@") Then xUpdate = True
'    If clpCode(5) = Format(rsEmp("ED_SECTION"), "@") Then xUpdate = True
'    If xUpdate = True Then
'        rsEmp("ED_DEPTNO") = clpDept
'        rsEmp("ED_DIV") = clpDiv
'        rsEmp("ED_GLNO") = clpGLNum
'        rsEmp("ED_EMP") = clpCode(4)
'        rsEmp("ED_ORG") = clpCode(0)
'        rsEmp("ED_PT") = clpPT
'        rsEmp("ED_EMPTYPE") = txtEmpType
'        rsEmp("ED_PAYROLL_ID") = txtPayrollID
'        rsEmp("ED_SECTION") = clpCode(5)
'    End If
'
'End Sub

'Private Sub UpdOttawaCCAC()
'    Dim rsEmp As New ADODB.Recordset
'    Dim rsTA As New ADODB.Recordset
'
'    rsEmp.Open "SELECT ED_EMPNBR,ED_DEPTNO,ED_DEPTEDATE,ED_DIVEDATE,ED_DIV,ED_GLNO,ED_PT,ED_EMPTYPE,ED_ORG,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If rsEmp.EOF Then Exit Sub
'
'    If (clpDept <> rsEmp("ED_DEPTNO") And Len(clpDept) > 0) Or (clpGLNum <> rsEmp("ED_GLNO") And Len(clpGLNum) > 0) Then
'        rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'        rsTA.AddNew
'        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN": rsTA("AU_UPLOAD") = "N"
'        rsTA("AU_TYPE") = "M"
'        rsTA("AU_NEWEMP") = "N"
'        If (clpDept <> rsEmp("ED_DEPTNO") And Len(clpDept) > 0) Then
'            rsTA("AU_OLDDEPT") = rsEmp("ED_DEPTNO")
'            rsTA("AU_DEPTNO") = clpDept
'        End If
'        If rsEmp("ED_GLNO") <> clpGLNum Then
'            If clpGLNum.Text <> "" Then
'                rsTA("AU_DEPT_GL") = clpGLNum.Text
'            Else
'                rsTA("AU_DEPT_GL") = Null
'            End If
'        End If
'        rsTA("AU_COMPNO") = "001"
'        rsTA("AU_EMPNBR") = glbLEE_ID
'        rsTA("AU_LDATE") = Date
'        rsTA("AU_LUSER") = glbUserID
'        rsTA("AU_LTIME") = Time$
'        rsTA("AU_UPLOAD") = "N"
'        rsTA.Update
'        rsTA.Close
'    End If
'    If Len(clpDept) > 0 Then
'        rsEmp("ED_DEPTNO") = clpDept
'        If clpDept <> rsEmp("ED_DEPTNO") Then
'            rsEmp("ED_DEPTEDATE") = dlpStartDate
'        End If
'    End If
'
'    If Len(clpDiv) > 0 Then
'        rsEmp("ED_DIV") = clpDiv
'        If clpDiv <> rsEmp("ED_DIV") Then
'            rsEmp("ED_DIVEDATE") = dlpStartDate
'        End If
'    End If
'
'    If Len(clpGLNum) > 0 Then rsEmp("ED_GLNO") = clpGLNum
'    If Len(clpCode(4)) > 0 Then rsEmp("ED_EMP") = clpCode(4)
'    If Len(clpCode(0)) > 0 Then rsEmp("ED_ORG") = clpCode(0)
'    If Len(clpPT) > 0 Then rsEmp("ED_PT") = clpPT
'    If Len(txtEmpType) > 0 Then rsEmp("ED_EMPTYPE") = txtEmpType
'
'    rsEmp("ED_EMPNBR") = glbLEE_ID
'    rsEmp.Update
'End Sub
'Ticket #16189-------------------------------

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
        xID = Data1.Recordset!TW_ID
    End If
    
    chkBenefitPayID = False
    rsTemp.Open "SELECT TW_ID FROM HR_TEMP_WORK WHERE TW_CURRENT<>0 AND TW_USRCHECK<>0 AND TW_EMPNBR=" & glbLEE_ID & " AND TW_ID<>" & xID & " AND TW_PAYROLL_ID<>'" & txtPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not rsTemp.EOF Then
        chkBenefitPayID = True
    End If
    
End Function

'Ticket #16189-------------------------------
'Private Function GetDoh(xEmpNo)
'Dim rs As New ADODB.Recordset
'Dim SQLQ
'    GetDoh = ""
'    SQLQ = "SELECT ED_EMPNBR,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
'    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rs.EOF Then
'        GetDoh = rs("ED_DOH")
'    End If
'    rs.Close
'End Function

'Private Sub updateOMERS()
'    'added by Bryan for Timmins 22/sep/05 Ticket#9368
'    Dim retVal As String
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT ED_EMPNBR, ED_DOB, ED_DEPTNO, ED_OMERS, ED_DOH, ED_NORMALR FROM HREMP "
'    strSQL = strSQL & "WHERE ED_EMPNBR = " & glbLEE_ID
'    rs.Open strSQL, gdbAdoIhr001, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If rs.EOF = False And rs.BOF = False Then
'        If chkCurrent(0).Value = True Then
'            Select Case clpPayrollCategory.Text
'            Case "001", "002", "003", "004", "005"
'                If rs("ED_DEPTNO") = "1510" Or rs("ED_DEPTNO") = "1600" Then
'                   'fire and police departments retire at 60
'                   If Not IsNull(rs("ED_DOB")) Then
'                        retVal = DateAdd("yyyy", 60, rs("ED_DOB"))
'                    End If
'                Else
'                    'the rest retire at 65
'                    If Not IsNull(rs("ED_DOB")) Then
'                        retVal = DateAdd("yyyy", 65, rs("ED_DOB"))
'                    End If
'                End If
'                rs("ED_NORMALR") = retVal
'                rs.Update
'            End Select
'        End If
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Sub

'Private Function fgetSection(xID) As String
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'    Dim retVal As String
'
'    If glbtermopen Then
'        strSQL = "SELECT ED_SECTION FROM TERM_HREMP WHERE TERM_SEQ =" & xID
'        rs.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
'    Else
'        strSQL = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & xID
'        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'    End If
'
'    If rs.EOF = False Then
'        If Not IsNull(rs("ED_SECTION")) Then
'            retVal = rs("ED_SECTION")
'        Else
'            retVal = ""
'        End If
'    Else
'        retVal = ""
'    End If
'    rs.Close
'    Set rs = Nothing
'
'    fgetSection = retVal
'
'End Function
'
'Private Function Get_DayHours_for_Job(xJob)
'    Dim rsHrJob As New ADODB.Recordset
'    Dim strSQL As String
'
'    strSQL = "SELECT JB_CODE, JB_DHRS FROM HRJOB WHERE JB_CODE = '" & xJob & "'"
'    rsHrJob.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    If Not rsHrJob.EOF Then
'        If Not IsNull(rsHrJob("JB_DHRS")) And rsHrJob("JB_DHRS") <> "" Then
'            Get_DayHours_for_Job = rsHrJob("JB_DHRS")
'        Else
'            Get_DayHours_for_Job = ""
'        End If
'    End If
'    rsHrJob.Close
'
'End Function
'Ticket #16189-------------------------------

Private Sub Update_Employee_Job_Training_List(xJob, xPosType, Optional xStartEndDate, Optional xEndDate)
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
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
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
            If flgUnqForPos Then
                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                'If xPosType = "Previous" And chkTrackCrsRenewal And chkCurrent(0) Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'T'"
                'Else
                '    If chkTrackCrsRenewal And chkCurrent(0) Then
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                '    Else
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & IIf(Left(xPosType, 1) = "C", "T", Left(xPosType, 1)) & "'"
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
                If flgUnqForPos Then
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
                'if the Course is Temporary and procedure not called from another procdure then
                'check if this course is required by another Primary "Current" Position if so then
                'change the xJob to that Position and Start Date to that Position Start Date
                If flgProcCalled = False And xPosType = "Temporary" Then
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND JH_CURRENT <> 0"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'Primacy Current Position requires this course so assign the Training List to Primary Current Job
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        xJob = rsEmpJob("TW_JOB")
                        xStartEndDate = rsEmpJob("TW_SDATE")
                        xPosType = "Current"
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                ElseIf flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (TW_ID <> " & rsDATA!TW_ID & ")"
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'Primary Current Position requires this course so assign the Training List to Primary Current Job
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
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                Else
                                    If xJob <> rsEmpJob("TW_JOB") Then   'If Temporary becoming Previous
                                        xPosType = "Temporary"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                xPosType = "Current"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            Else
                                If xJob <> rsEmpJob("TW_JOB") Then   'If Temporary becoming Previous
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        End If
                    Else
                        xStartEndDate = ""      'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
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
                                If IsMissing(xEndDate) Or xEndDate = "" Then
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
                    'Ticket #24300
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
                    
                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                    rsHRTrain.Update
                    
'                    rsFollowUp.Close
'                    Set rsFollowUp = Nothing
                    
                    'Update Temp/Cross Training Position record with Follow Up ID
                    'if the course code is TRAIN
                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                        'Search HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                'if the Course is Temporary and procedure not called from another procdure then
                'check if this course is required by another Primary "Current" Position if so then
                'change the xJob to that Position and Start Date to that Position Start Date
                If flgProcCalled = False And xPosType = "Temporary" Then
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND JH_CURRENT <> 0"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'Primacy Current Position requires this course so assign the Training List to Primary Current Job
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        xJob = rsEmpJob("TW_JOB")
                        xStartEndDate = rsEmpJob("TW_SDATE")
                        xPosType = "Current"
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                ElseIf flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (TW_ID <> " & rsDATA!TW_ID & ")"
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'Primacy Current Position requires this course so assign the Training List to Primary Current Job
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
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                Else
                                    If xJob <> rsEmpJob("TW_JOB") Then   'If Temporary becoming Previous
                                        xPosType = "Temporary"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                xPosType = "Current"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            Else
                                If xJob <> rsEmpJob("TW_JOB") Then   'If Temporary becoming Previous
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        End If
                    Else
                        xStartEndDate = ""      'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                End If
                
                
                
                
                
                'Training record for this course already exists so update the Renewal Date
                'Check which Type of Position is assigned to this course
                If rsHRTrain("TR_POS_TYPE") = "C" Then
                    'Currently the course is holding Primary Current Position Code
                    'This is Temporary Position - so no change is to be made according to design
                    'Do not do anything
                ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                    'Currently the course is holding Temporary Current Position
                    'Check which type of position requires this course
                    'Changed 2
                    'If xPosType = "Current" Then
                    If xPosType = "Temporary" Then
                        'This course is for new Current Temp Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
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
                                rsHRTrain("TR_POS_TYPE") = "T"   'Current Temporary
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
                                    'Ticket #24300
                                    'Add a Follow Up record for this Training course
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                Else
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
                                
                                rsHRTrain.Update
                                
                                'Update Temp/Cross Training Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                
                                    'Clear the Follow Up Id on the other current Temp position rec
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
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                        
                                        'Clear the Follow Up Id on the Temp/Cross Training Position record
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
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                    
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
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
                        Else
                            'Do not do anything because Training List has most recent Position Start Date
                        End If
                    ElseIf xPosType = "Previous" Then
                        'TEMPORARY - Previous
                        'Temporary Position becoming Previous
                        'Current Temp Position is holding this course but Previous Temp Position requires this course
                        'Check if the Position in HR_TRAIN is same this Position
                        If (rsHRTrain("TR_JOB") <> xJob) Or (rsHRTrain("TR_JOB") = xJob And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(dlpStartDate.Text) And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate))) Then
                            'Do not do anything because Current takes the  priority
                        Else
                            'Renewal Date based on last Course Taken date if present
                            'otherwise Follow Up Effective Date Period
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
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
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Temporary
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
                                    'Ticket #24300
                                    'Add a Follow Up record for this Training course
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                Else
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
                                
                                rsHRTrain.Update
                                
                                'Update Temp/Cross Training Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                        
                                        'Clear the Follow Up ID in the Temp/Cross Training Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                    
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                    'Previous Primary or Temporary position is holding this course
                    'Changed 2
                    'If xPosType = "Current" Then
                    If xPosType = "Temporary" Then
                        'This course is for new Current Temp Position so recalculate the
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
                            rsHRTrain("TR_POS_TYPE") = "T"   'Current Temporary
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
                                'Ticket #24300
                                'Add a Follow Up record for this Training course
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                            Else
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
                        
                            rsHRTrain.Update
                            
                            'Update Temp/Cross Training Position record with Follow Up ID
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
                                
                                'Search HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                        
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
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
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
                        'These courses are for new Previous Temp Position so recalculate the
                        'Renewal Dates
                        xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                        If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                        'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate), dlpStartDate.Text, xStartEndDate)) Then
                        If CVDate(xPrvEndDate) < CVDate(IIf(IsMissing(xEndDate) Or xEndDate = "", dlpENDDATE.Text, xEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
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
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Temporary
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
                                    'Ticket #24300
                                    'Add a Follow Up record for this Training course
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                Else
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
                            
                                rsHRTrain.Update
                                
                                'Update Temp/Cross Training Position record with Follow Up ID
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
                                    
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
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
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
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
                ElseIf IsNull(rsHRTrain("TR_POS_TYPE")) Or rsHRTrain("TR_POS_TYPE") = "" Then
                    'Check if the course was taken before. If taken then use the normal Training List logic bases
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
                        
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            'Ticket #24300
                            'Add a Follow Up record for this Training course
                            rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                        Else
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
                        
                        rsHRTrain.Update
                        
                        'Update Temp/Cross Training Position record with Follow Up ID
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
                            
                            'Search HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                            'Renewal Period found - update existing records
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
                                                    
                            'Renewal period available
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
                            
                            
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Ticket #24300
                                'Add a Follow Up record for this Training course
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                            Else
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
                            
                            rsHRTrain.Update
                            
                            'Update Temp/Cross Training Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Search HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'No Job - Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
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
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                    
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
                            SQLQ = SQLQ & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
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
    
    'Ticket #22951
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
                            
                            'Clear the Follow Up ID in the Temp/Cross Training Position record
                            'if the course code is TRAIN
                            If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                                'Search HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
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
                
                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                    'if the course code is TRAIN
                    If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                        'Search HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("TW_FOLLOWUP_ID") = Null
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
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                'and not the position currently selected
                SQLQ = SQLQ & " AND (TW_ID <> " & rsDATA!TW_ID & ")"
                rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmpJob.EOF Then
                    rsEmpJob.MoveFirst
                    
                    Do While Not rsEmpJob.EOF
                        'Check in the Required Courses table if the retrieved required course is required by othe retrieved position
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
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        
                            'Clear the Follow Up ID in the Temp/Cross Training Position record
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Search HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                        gdbAdoIhr001.Execute SQLQ
                        
                        'Clear the Follow Up Id in the Temp/Cross Training Position record
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Search HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & Data1.Recordset("TW_ID")
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
    Dim SQLQ, xDWMY As String
    Dim xRenewalDt
    Dim xComments As String

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
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
            SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            'and not the position currently selected
            SQLQ = SQLQ & " AND (TW_ID <> " & rsDATA!TW_ID & ")"
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
                                    
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Ticket #24300
                                        'Add a Follow Up record for this Training course
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), rsEmpJobs("TW_JOB"))
                                    Else
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
                                    
                                    rsHRTrain.Update
                                    
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
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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
                                    
                                    
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Ticket #24300
                                        'Add a Follow Up record for this Training course
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsCourseCode("ES_CRSCODE"), rsEmpJobs("TW_JOB"))
                                    Else
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
                                
                                    rsHRTrain.Update
                                    
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
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
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

End Sub

Private Sub Update_Position_Start_Date_in_Training_List(oldSDate, newSDate)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ, xDWMY As String
    Dim xComments As String
    
    'Retrieve Training List records which match this employee, job and original start date
    SQLQ = "SELECT * FROM HR_TRAIN "
    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
    SQLQ = SQLQ & " AND TR_SDATE = " & Date_SQL(oldSDate)
    SQLQ = SQLQ & " AND TR_POS_TYPE <> 'C'"
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
                                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Ticket #24300
                    'Add a Follow Up record for this Training course
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

Private Function PrvSDate(xCurrent)
Dim SQLQ As String
Dim HRJH_Snap As New ADODB.Recordset

PrvSDate = 0    ' returns 0 if no found records

On Error GoTo PrvSDate_Err

SQLQ = "Select TW_EMPNBR, TW_SDATE FROM HR_TEMP_WORK"
SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " "
If xCurrent Then
    SQLQ = SQLQ & " AND TW_CURRENT =0"
End If
SQLQ = SQLQ & " ORDER BY TW_SDATE DESC"
HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRJH_Snap.BOF And HRJH_Snap.EOF Then
    Exit Function
Else
    PrvSDate = HRJH_Snap("TW_SDATE")
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

Private Function Get_Current_Primary_Job()
    Dim rsJobHist As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_Current_Primary_Job = ""
    SQLQ = "SELECT JH_EMPNBR,JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID & " AND JH_CURRENT <> 0"
    rsJobHist.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJobHist.EOF Then
        Get_Current_Primary_Job = rsJobHist("JH_JOB")
    Else
        Get_Current_Primary_Job = ""
    End If
    rsJobHist.Close
    Set rsJobHist = Nothing
End Function
