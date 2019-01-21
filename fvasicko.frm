VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmVACSICKO 
   Appearance      =   0  'Flat
   Caption         =   "Vacation and Sick Overview"
   ClientHeight    =   9285
   ClientLeft      =   -45
   ClientTop       =   1365
   ClientWidth     =   12825
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
   ScaleHeight     =   9285
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fvasicko.frx":0000
      Height          =   3195
      Left            =   0
      OleObjectBlob   =   "fvasicko.frx":0014
      TabIndex        =   0
      Tag             =   "Employee Listing "
      Top             =   120
      Visible         =   0   'False
      Width           =   11685
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
      Bindings        =   "fvasicko.frx":889C
      Height          =   3195
      Left            =   0
      OleObjectBlob   =   "fvasicko.frx":88B0
      TabIndex        =   40
      Tag             =   "Employee Listing "
      Top             =   120
      Width           =   11655
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGridVL 
      Height          =   3135
      Left            =   12360
      OleObjectBlob   =   "fvasicko.frx":10FF5
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGridVLDay 
      Height          =   3135
      Left            =   11760
      OleObjectBlob   =   "fvasicko.frx":16547
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin Threed.SSPanel panCalcEnt 
      Height          =   1680
      Left            =   7485
      TabIndex        =   52
      Top             =   4335
      Visible         =   0   'False
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   2963
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
      Begin MSMask.MaskEdBox medCalcSick 
         DataField       =   "WK_CSDAY"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   53
         Tag             =   "11-Projected Year-End Accrual"
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   1215
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCalcSick 
         DataField       =   "ED_ANNSICK"
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   55
         TabStop         =   0   'False
         Tag             =   "11-Projected Year-End Accrual"
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   1215
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCalcVac 
         DataField       =   "ED_ANNVAC"
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   56
         TabStop         =   0   'False
         Tag             =   "11-Projected Year-End Accrual"
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   795
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medAvailSick 
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   88
         Tag             =   "11-Projected Year-End to Book"
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   1215
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medAvailVac 
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   90
         Tag             =   "11-Projected Year-End to Book"
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   795
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medAvailVac 
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   91
         TabStop         =   0   'False
         Tag             =   "11-Projected Year-End to Book"
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   795
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medAvailSick 
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   92
         TabStop         =   0   'False
         Tag             =   "11-Projected Year-End to Book"
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   1215
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medToBeSick 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   94
         Top             =   1215
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medToBeVac 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   95
         Top             =   795
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medToBeSick 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   96
         Top             =   1215
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medToBeVac 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   97
         Top             =   795
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCalcVac 
         DataField       =   "WK_CVDAY"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   105
         Tag             =   "11-Projected Year-End Accrual"
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   795
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "to Book"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   1620
         TabIndex        =   110
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   560
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accrual"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   360
         TabIndex        =   109
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   560
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year-End"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   1620
         TabIndex        =   107
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   340
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year-End"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   360
         TabIndex        =   106
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   340
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "*Projected"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   1620
         TabIndex        =   89
         ToolTipText     =   "Projected Year-End to Book"
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Projected"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   54
         ToolTipText     =   "Projected Year-End Accrual"
         Top             =   120
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   39
      Top             =   8625
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdReCompDAccrual 
         Appearance      =   0  'Flat
         Caption         =   "&Re-Create Daily Accrual"
         Height          =   375
         Left            =   9120
         TabIndex        =   111
         Tag             =   "Employee's Daily Accrual as of Entitlement Start Date"
         Top             =   60
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "Show Calculated Entitlements"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   18
         Top             =   60
         Width           =   2745
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate 1 Employee"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   16
         Tag             =   "Recalculate for the employee"
         Top             =   60
         Width           =   2415
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate All Employees"
         Height          =   375
         Index           =   1
         Left            =   11040
         TabIndex        =   17
         Tag             =   "Recalculate for all employees"
         Top             =   60
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton cmdHours 
         Appearance      =   0  'Flat
         Caption         =   "&Hours"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1260
         TabIndex        =   15
         Tag             =   "Display Vacation and Sick Overview in Hours"
         Top             =   60
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6000
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.CommandButton cmdDays 
         Appearance      =   0  'Flat
         Caption         =   "Da&ys"
         Height          =   375
         Left            =   300
         TabIndex        =   14
         Tag             =   "Display Vacation and Sick Overview in Days"
         Top             =   60
         Width           =   875
      End
   End
   Begin VB.CommandButton cmdModify1 
      Appearance      =   0  'Flat
      Caption         =   "&Edit"
      Height          =   375
      Left            =   10920
      TabIndex        =   41
      Tag             =   "Edit information on this screen"
      Top             =   5115
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdOK1 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   46
      Tag             =   "Save changes made"
      Top             =   5580
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel1 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   47
      Tag             =   "Cancel changes made"
      Top             =   6060
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   8280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Tag             =   "Find Employee"
      Top             =   6420
      Width           =   735
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Tag             =   "00-Search for Surname"
      Top             =   6450
      Width           =   1935
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Surname"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   13
      Top             =   6420
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Emp #"
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   19
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   6420
      Width           =   2475
   End
   Begin Threed.SSPanel panDetails 
      Height          =   2655
      Left            =   0
      TabIndex        =   25
      Top             =   3450
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   4683
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
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin MSMask.MaskEdBox medSickR 
         DataField       =   "WK_SICKODAY"
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Current Outstanding"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSickC 
         DataField       =   "WK_SICKTDAY"
         Height          =   285
         Index           =   1
         Left            =   3270
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Total Taken and Booked"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacC 
         DataField       =   "WK_VACTDAY"
         Height          =   285
         Index           =   1
         Left            =   3285
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Total Taken and Booked"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPSick 
         DataField       =   "WK_PSICKDAY"
         Height          =   285
         Index           =   1
         Left            =   930
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of sicktime from previous year"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPVac 
         DataField       =   "WK_PVACDAY"
         Height          =   285
         Index           =   1
         Left            =   930
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of vacation from previous year"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCSickDay 
         DataField       =   "WK_SICKDAY"
         Height          =   285
         Left            =   2100
         TabIndex        =   7
         Top             =   2100
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCVacDay 
         DataField       =   "WK_VACDAY"
         Height          =   285
         Left            =   2100
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSickR 
         DataField       =   "WK_SICKO"
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Current Outstanding"
         Top             =   2100
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSickC 
         DataField       =   "ED_SICKT"
         Height          =   285
         Index           =   0
         Left            =   3270
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Total Taken and Booked"
         Top             =   2100
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacC 
         DataField       =   "ED_VACT"
         Height          =   285
         Index           =   0
         Left            =   3285
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Total Taken and Booked"
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPSick 
         DataField       =   "ED_PSICK"
         Height          =   285
         Index           =   0
         Left            =   930
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of sicktime from previous year"
         Top             =   2100
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCSick 
         DataField       =   "ED_SICK"
         Height          =   285
         Left            =   2100
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "11-Total Number of hours of Sicktime entitled"
         Top             =   2100
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPVac 
         DataField       =   "ED_PVAC"
         Height          =   285
         Index           =   0
         Left            =   930
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of vacation from previous year"
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCVac 
         DataField       =   "ED_VAC"
         Height          =   285
         Left            =   2100
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "11-Total number of hours vacation time entitled"
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacR 
         DataField       =   "WK_VACODAY"
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Current Outstanding"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacR 
         DataField       =   "WK_VACO"
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Current Outstanding"
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.DateLookup dlpFDate1 
         DataField       =   "ED_EFDATE"
         Height          =   285
         Left            =   7560
         TabIndex        =   42
         Tag             =   "40-Vacation Start date"
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTDate1 
         DataField       =   "ED_ETDATE"
         Height          =   285
         Left            =   9240
         TabIndex        =   43
         Tag             =   "40-Vacation ending date"
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFDate1S 
         DataField       =   "ED_EFDATES"
         Height          =   285
         Left            =   7560
         TabIndex        =   44
         Tag             =   "40-Vacation Start date"
         Top             =   120
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTDate1S 
         DataField       =   "ED_ETDATES"
         Height          =   285
         Left            =   9240
         TabIndex        =   45
         Tag             =   "40-Vacation ending date"
         Top             =   120
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTmpEndDate 
         Height          =   285
         Left            =   360
         TabIndex        =   100
         Tag             =   "40-Temporary Vacation End date"
         Top             =   0
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1435
      End
      Begin MSMask.MaskEdBox medTmpVTakenHr 
         Height          =   285
         Left            =   720
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTmpVTakenDy 
         Height          =   285
         Left            =   720
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Booked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   3285
         TabIndex        =   108
         ToolTipText     =   "Total Taken and Booked"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Outstanding"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   4425
         TabIndex        =   104
         ToolTipText     =   "Current Outstanding"
         Top             =   1220
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taken &&"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   3285
         TabIndex        =   103
         ToolTipText     =   "Total Taken and Booked"
         Top             =   1220
         Width           =   1005
      End
      Begin VB.Label DateSeleV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ED_EFDATE"
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
         Height          =   285
         Index           =   1
         Left            =   7740
         TabIndex        =   51
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label DateSeleS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ED_EFDATES"
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
         Height          =   285
         Index           =   1
         Left            =   7740
         TabIndex        =   50
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label DateSeleV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ED_ETDATE"
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
         Height          =   285
         Index           =   2
         Left            =   9000
         TabIndex        =   49
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label DateSeleS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ED_ETDATES"
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
         Height          =   285
         Index           =   2
         Left            =   9000
         TabIndex        =   48
         Top             =   2100
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   7740
         TabIndex        =   37
         Top             =   1215
         Width           =   2235
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Based On"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   5610
         TabIndex        =   36
         Top             =   1220
         Width           =   1815
      End
      Begin VB.Label DateSeleV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entitlements Date"
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
         Height          =   285
         Index           =   0
         Left            =   5610
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label DateSeleS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entitlements Date"
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
         Height          =   285
         Index           =   0
         Left            =   5625
         TabIndex        =   10
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label lblDayHrs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HOURS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4425
         TabIndex        =   30
         ToolTipText     =   "Current Outstanding"
         Top             =   1005
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3285
         TabIndex        =   29
         ToolTipText     =   "Total Taken and Booked"
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   23
         Top             =   1220
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sicktime"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   24
         Top             =   2145
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   28
         Top             =   1220
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   26
         Top             =   1725
         Width           =   765
      End
   End
   Begin Threed.SSCheck chkEndDateEdit 
      Height          =   255
      Left            =   11280
      TabIndex        =   98
      Tag             =   "Temporary Vacation End Date to calculate Vacaton Taken"
      Top             =   5115
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel panVADetails 
      Height          =   1575
      Left            =   120
      TabIndex        =   59
      Top             =   6960
      Visible         =   0   'False
      Width           =   6915
      _Version        =   65536
      _ExtentX        =   12197
      _ExtentY        =   2778
      _StockProps     =   15
      BackColor       =   14215660
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
      Begin MSMask.MaskEdBox medCalcVac 
         DataField       =   "WK_CVDAY"
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   60
         Top             =   435
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCalcVac 
         DataField       =   "ED_ANNVAC"
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   61
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of vacation from previous year"
         Top             =   435
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel panVCurrent 
         Height          =   1335
         Left            =   960
         TabIndex        =   66
         Top             =   120
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   2355
         _StockProps     =   15
         BackColor       =   14215660
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
         Begin MSMask.MaskEdBox medSickR 
            DataField       =   "WK_SICKODAY"
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   975
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medSickR 
            DataField       =   "WK_SICKO"
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   975
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medVacR 
            DataField       =   "WK_VACODAY"
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   315
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medVacR 
            DataField       =   "WK_VACO"
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   315
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Accrued"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   630
         End
      End
      Begin Threed.SSPanel panVOutst 
         Height          =   1335
         Left            =   5640
         TabIndex        =   73
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   2355
         _StockProps     =   15
         BackColor       =   14215660
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
         Enabled         =   0   'False
         Begin MSMask.MaskEdBox medSickRVLDay 
            Height          =   285
            Left            =   0
            TabIndex        =   74
            Top             =   975
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medVacRVLDay 
            Height          =   285
            Left            =   0
            TabIndex        =   75
            Top             =   315
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medSickRVL 
            DataField       =   "ED_SICK"
            Height          =   285
            Left            =   0
            TabIndex        =   76
            TabStop         =   0   'False
            Tag             =   "11-Total Number of hours of Sicktime entitled"
            Top             =   975
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medVacRVL 
            Height          =   285
            Left            =   0
            TabIndex        =   77
            TabStop         =   0   'False
            Tag             =   "11-Total number of hours vacation time entitled"
            Top             =   315
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   78
            Top             =   720
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin Threed.SSPanel ssVSickT 
         Height          =   615
         Left            =   2760
         TabIndex        =   80
         Top             =   840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   15
         BackColor       =   14215660
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
         Enabled         =   0   'False
         Begin MSMask.MaskEdBox medCalcSick 
            DataField       =   "WK_CSDAY"
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   81
            Top             =   255
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medCalcSick 
            DataField       =   "ED_ANNSICK"
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   82
            TabStop         =   0   'False
            Tag             =   "11-Banked hours of sicktime from previous year"
            Top             =   255
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Taken"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   555
         End
      End
      Begin Threed.SSPanel panVTaken 
         Height          =   735
         Left            =   4320
         TabIndex        =   84
         Top             =   110
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1296
         _StockProps     =   15
         BackColor       =   14215660
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
         Enabled         =   0   'False
         Begin MSMask.MaskEdBox medVacC 
            DataField       =   "WK_VACTDAY"
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   340
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medVacC 
            DataField       =   "ED_VACT"
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   340
            Width           =   1005
            _ExtentX        =   1773
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
            Format          =   "Fixed"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Taken"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   87
            Top             =   10
            Width           =   735
         End
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual Entitlement"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   2640
         TabIndex        =   64
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sicktime"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   0
         TabIndex        =   63
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   62
         Top             =   480
         Width           =   765
      End
      Begin VB.Label lblVLonly 
         Caption         =   "VitalAire only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   10920
      TabIndex        =   99
      Top             =   4695
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* - (Previous + Calculated) - Taken"
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
      Index           =   19
      Left            =   120
      TabIndex        =   93
      Top             =   8880
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   38
      Top             =   6480
      Width           =   1665
   End
End
Attribute VB_Name = "frmVACSICKO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add a New True DBGRid
Option Explicit
Dim EESNameSort As Integer
Dim OSN As Double, OSCh As String     ' last search items
Dim fglbWDate$, SavEntOpt, SavFdate, SavTdate
Dim fglbWDateS$, SavEntOptS, SAVFDATES, SAVTDATES
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsEntRules As New ADODB.Recordset
Dim fglbNew As Integer
Dim fglbESQLQ As String, snapEntitle As New ADODB.Recordset, fglbAsOf As Date
Dim fbtnCalc As Boolean

Private Sub chkEndDateEdit_Click(Value As Integer)
    If chkEndDateEdit.Value = True Then
        panDetails.Enabled = True
        
        cmdOK1.Enabled = True
        cmdCancel1.Enabled = True
        
        panCalcEnt.Visible = False
        If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
            cmdCalc.Caption = "Show To Be Accrued"
        ElseIf glbCompSerial = "S/N - 2430W" Then
            'Ticket #23878 - Carizon Family and Community Services/KidsLink
            cmdCalc.Caption = "Show Annual Entitlements"
        Else
            cmdCalc.Caption = "Show Calculated Entitlements"
        End If
        
        dlpTmpEndDate.Visible = True
        DateSeleV(2).Visible = False
        dlpTmpEndDate.Top = DateSeleV(2).Top
        dlpTmpEndDate.Left = DateSeleV(2).Left
        medTmpVTakenHr.Top = medVacC(1).Top
        medTmpVTakenHr.Left = medVacC(1).Left
        medTmpVTakenDy.Top = medVacC(1).Top
        medTmpVTakenDy.Left = medVacC(1).Left
        
        If lblDayHrs.Caption = "DAYS" Then
            medTmpVTakenDy.Visible = True
            medTmpVTakenHr.Visible = False
            medVacC(1).Visible = False
        Else
            medTmpVTakenDy.Visible = False
            medTmpVTakenHr.Visible = True
            medVacC(0).Visible = False
        End If
    Else
        panDetails.Enabled = False
        
        cmdOK1.Enabled = False
        cmdCancel1.Enabled = False
        cmdModify1.Enabled = True
        
        dlpTmpEndDate = ""
        dlpTmpEndDate.Visible = False
        DateSeleV(2).Visible = True
        
        medTmpVTakenDy = ""
        medTmpVTakenHr = ""
        medTmpVTakenDy.Visible = False
        medTmpVTakenHr.Visible = False
        
        If lblDayHrs.Caption = "DAYS" Then
            medVacC(1).Visible = True
        Else
            medVacC(0).Visible = True
        End If
    End If
    
End Sub

Private Sub cmdCalc_Click()
    'Ticket #23878 - Carizon Family and Community Services/KidsLink
    If cmdCalc.Caption = "Show Calculated Entitlements" Or cmdCalc.Caption = "Show Annual Entitlements" Then
        panCalcEnt.Visible = True
        cmdCalc.Caption = "Show Entitlement Dates"
        lblTitle(19).Visible = True
        
        'Ticket #23878 - KidsLink/Carizon - Available = Outstanding
        If glbCompSerial = "S/N - 2430W" Then
            lblTitle(19).Caption = "* Outstanding"
        End If
        
        '7.9 Enhancement - only allow "Show Calculated" to be clicked if Monthly Entitlement Update
        If glbCompEntVac$ <> "M" And glbCompEntVac$ <> "D" Then
            medCalcVac(0).Visible = False
            medCalcVac(1).Visible = False
            medAvailVac(0).Visible = False
            medAvailVac(1).Visible = False
        End If
        If glbCompEntSick$ <> "M" Then
            medCalcSick(0).Visible = False
            medCalcSick(1).Visible = False
            medAvailSick(0).Visible = False
            medAvailSick(1).Visible = False
        End If
        
        Call ReCalcAnn
        
    ElseIf cmdCalc.Caption = "Show To Be Accrued" Then
        panCalcEnt.Visible = True
        cmdCalc.Caption = "Show Entitlement Dates"
        
        medCalcVac(0).Visible = False
        medCalcSick(0).Visible = False
        medCalcVac(1).Visible = False
        medCalcSick(1).Visible = False
        
        If lblDayHrs.Caption = "DAYS" Then
            medToBeVac(0).Visible = True
            medToBeSick(0).Visible = True
            medToBeVac(1).Visible = False
            medToBeSick(1).Visible = False
        Else
            medToBeVac(0).Visible = False
            medToBeSick(0).Visible = False
            medToBeVac(1).Visible = True
            medToBeSick(1).Visible = True
        End If
        
        Call ReCalcAnn
        Call Calc_To_Be_Accrued
    Else
        panCalcEnt.Visible = False
        If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
            cmdCalc.Caption = "Show To Be Accrued"
        ElseIf glbCompSerial = "S/N - 2430W" Then
            'Ticket #23878 - Carizon Family and Community Services/KidsLink
            cmdCalc.Caption = "Show Annual Entitlements"
        Else
            cmdCalc.Caption = "Show Calculated Entitlements"
        End If
        lblTitle(19).Visible = False
    End If
    
End Sub

Private Sub cmdCancel1_Click()
Dim x, xID
On Error GoTo Can_Err

xID = Data1.Recordset("ED_EMPNBR")

rsDATA.CancelUpdate
Call Display_Value
Data1.Refresh

Data1.Recordset.Find "ED_EMPNBR=" & xID
    
panDetails.Enabled = False
cmdOK1.Enabled = False
cmdCancel1.Enabled = False
cmdModify1.Enabled = True
dlpFDate1.Visible = False
dlpTDate1.Visible = False
DateSeleV(1).Visible = True
DateSeleV(2).Visible = True

If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then  'City of Kawartha Lakes or Crown Investment Corp (Ticket #14084)
    dlpFDate1S.Visible = False
    dlpTDate1S.Visible = False
    DateSeleS(1).Visible = True
    DateSeleS(2).Visible = True
End If

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdDays_Click()
cmdDays.Enabled = False
cmdHours.Enabled = True
medCVac.Visible = False
medCVacDay.Visible = True
medPVac(0).Visible = False
medPVac(1).Visible = True
medVacC(0).Visible = False
medVacC(1).Visible = True
medVacR(0).Visible = False
medVacR(1).Visible = True
medCalcVac(0).Visible = True
medCalcVac(1).Visible = False
medAvailVac(0).Visible = True
medAvailVac(1).Visible = False

If Not glbWFC And glbCompSerial <> "S/N - 2418W" Then 'No Sick on this screen
    medCSick.Visible = False
    medCSickDay.Visible = True
    medPSick(0).Visible = False
    medPSick(1).Visible = True
    medSickC(0).Visible = False
    medSickC(1).Visible = True
    medSickR(0).Visible = False
    medSickR(1).Visible = True
    medCalcSick(0).Visible = True
    medCalcSick(1).Visible = False
    medAvailSick(0).Visible = True
    medAvailSick(1).Visible = False
    
    '7.9 Enhancement - only allow "Show Calculated" to be clicked if Monthly Entitlement Update
    If glbCompSerial <> "S/N - 2395W" Then
        If glbCompEntVac$ <> "M" And glbCompEntVac$ <> "D" Then  'Or glbCompEntSick$ <> "M" Then
            medCalcVac(0).Visible = False
            medCalcVac(1).Visible = False
            medAvailVac(0).Visible = False
            medAvailVac(1).Visible = False
        End If
        If glbCompEntSick$ <> "M" Then
            medCalcSick(0).Visible = False
            medCalcSick(1).Visible = False
            medAvailSick(0).Visible = False
            medAvailSick(1).Visible = False
        End If
    End If
    
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
        
        medToBeVac(0).Visible = True
        medToBeSick(0).Visible = True
        medToBeVac(1).Visible = False
        medToBeSick(1).Visible = False
        
        medCalcVac(0).Visible = False
        medCalcSick(0).Visible = False
        medCalcVac(1).Visible = False
        medCalcSick(1).Visible = False
    End If
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    medVacR(2).Visible = Not False
    medVacR(3).Visible = Not True
    medCalcVac(2).Visible = Not False
    medCalcVac(3).Visible = Not True
    
    If Not GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        lblTitle(11).Visible = False
        lblTitle(12).Visible = False
        medCalcVac(2).Enabled = False
        medCalcVac(3).Enabled = False
        panVCurrent.Enabled = False 'Ticket #14169
        medVacRVLDay.Visible = False
        medVacRVL.Visible = False
        medSickRVLDay.Visible = False
        medSickRVL.Visible = False
    Else
        lblTitle(11).Visible = True
        lblTitle(12).Visible = True
        medCalcVac(2).Enabled = True
        medCalcVac(3).Enabled = True
        panVCurrent.Enabled = True 'Ticket #14169
        medVacRVLDay.Visible = Not False
        medVacRVL.Visible = Not True
        medSickRVLDay.Visible = Not False
        medSickRVL.Visible = Not True
    End If
    medSickR(2).Visible = Not False
    medSickR(3).Visible = Not True
    medCalcSick(2).Visible = Not False
    medCalcSick(3).Visible = Not True
    vbxTrueGridVL.Visible = False
    vbxTrueGridVLDay.Visible = True
    vbxTrueGrid.Visible = False
    vbxTrueGrid1.Visible = False
Else
    vbxTrueGrid.Visible = True
    vbxTrueGrid1.Visible = False
    
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Days_Data1_Source
    End If
End If

lblDayHrs.Caption = "DAYS"

If chkEndDateEdit.Value = True Then
    Call chkEndDateEdit_Click(True)
    Call dlpTmpEndDate_LostFocus
End If

End Sub

Private Sub cmdDays_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdEESort_Click(Index As Integer)

txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort(0).Visible = False
    cmdEESort(1).Visible = True
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort(0).Visible = True
    cmdEESort(1).Visible = False
End If

If EERetrieve() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = " "
txtEESearch.SetFocus

End Sub

Private Sub cmdEESort_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark

On Error GoTo Srch_Err

If Not Len(txtEESearch) > 0 Then
   MsgBox "To search you must enter something to search for."
   Exit Sub
End If
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch.Text, "'", "''")
    If EESNameSort = True Then
        SQLQ = "ED_SURNAME  >= '" & Sch & "'"
    Else
        If Not IsNumeric(txtEESearch.Text) And Not glbLinamar Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            SQLQ = "EMPNBR >= '" & Sch & "'"
        Else
            SQLQ = "ED_EMPNBR >= '" & Sch & "'"
        End If

    End If
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    MsgBox "Employee not found"
    Data1.Refresh
End If

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "Find Next")
Call RollBack '28July99 jsEnd Sub
End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdHours_Click()

cmdDays.Enabled = True
cmdHours.Enabled = False
medCVac.Visible = True
medCVacDay.Visible = False
medPVac(0).Visible = True
medPVac(1).Visible = False
medVacC(0).Visible = True
medVacC(1).Visible = False
medVacR(0).Visible = True
medVacR(1).Visible = False
medCalcVac(0).Visible = False
medCalcVac(1).Visible = True
medAvailVac(0).Visible = False
medAvailVac(1).Visible = True

If Not glbWFC And glbCompSerial <> "S/N - 2418W" Then
    medCSick.Visible = True
    medCSickDay.Visible = False
    medPSick(0).Visible = True
    medPSick(1).Visible = False
    medSickC(0).Visible = True
    medSickC(1).Visible = False
    medSickR(0).Visible = True
    medSickR(1).Visible = False
    medCalcSick(0).Visible = False
    medCalcSick(1).Visible = True
    medAvailSick(0).Visible = False
    medAvailSick(1).Visible = True

    '7.9 Enhancement - only allow "Show Calculated" to be clicked if Monthly Entitlement Update
    If glbCompSerial <> "S/N - 2395W" Then
        If glbCompEntVac$ <> "M" And glbCompEntVac$ <> "D" Then 'Or glbCompEntSick$ <> "M" Then
            medCalcVac(0).Visible = False
            medCalcVac(1).Visible = False
            medAvailVac(0).Visible = False
            medAvailVac(1).Visible = False
        End If
        If glbCompEntSick$ <> "M" Then
            medCalcSick(0).Visible = False
            medCalcSick(1).Visible = False
            medAvailSick(0).Visible = False
            medAvailSick(1).Visible = False
        End If
    End If

    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
        
        medToBeVac(0).Visible = False
        medToBeSick(0).Visible = False
        medToBeVac(1).Visible = True
        medToBeSick(1).Visible = True
        
        medCalcVac(0).Visible = False
        medCalcSick(0).Visible = False
        medCalcVac(1).Visible = False
        medCalcSick(1).Visible = False
    End If
End If
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    medVacR(2).Visible = False
    medVacR(3).Visible = True
    medCalcVac(2).Visible = False
    medCalcVac(3).Visible = True
    
    If Not GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        lblTitle(11).Visible = False
        lblTitle(12).Visible = False
        medCalcVac(2).Enabled = False
        medCalcVac(3).Enabled = False
        panVCurrent.Enabled = False 'Ticket #14169
        medVacRVLDay.Visible = False
        medVacRVL.Visible = False
        medSickRVLDay.Visible = False
        medSickRVL.Visible = False
    Else
        lblTitle(11).Visible = True
        lblTitle(12).Visible = True
        medCalcVac(2).Enabled = True
        medCalcVac(3).Enabled = True
        panVCurrent.Enabled = True 'Ticket #14169
        medVacRVLDay.Visible = False
        medVacRVL.Visible = True
        medSickRVLDay.Visible = False
        medSickRVL.Visible = True
    End If
    medSickR(2).Visible = False
    medSickR(3).Visible = True
    medCalcSick(2).Visible = False
    medCalcSick(3).Visible = True
    vbxTrueGridVL.Visible = True
    vbxTrueGridVLDay.Visible = False
    vbxTrueGrid.Visible = False
    vbxTrueGrid1.Visible = False
Else
    vbxTrueGrid.Visible = False
    vbxTrueGrid1.Visible = True
    
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Hours_Data1_Source
    End If
End If

lblDayHrs.Caption = "HOURS"

If chkEndDateEdit.Value = True Then
    Call chkEndDateEdit_Click(True)
    Call dlpTmpEndDate_LostFocus
End If

End Sub

Private Sub cmdHours_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify1_Click()
Dim xMsg
    If glbCompSerial <> "S/N - 2363W" And glbCompSerial <> "S/N - 2205W" Then    'City of Kawartha Lakes or Crown Investment Corp (Ticket #14084)
        'Ticket #30142 - Opening up to edit the FT employee's entitlement and date range as well
        'If Not (Data1.Recordset("ED_PT") = "PT" And Data1.Recordset("ED_ORG") = "CUPE") Then
        '    xMsg = "You only can edit the Vacation Date Range" & Chr(10)
        '    xMsg = xMsg & "when the employee type is Part Time and union is 'CUPE' "
        '    MsgBox xMsg
        '    Exit Sub
        'End If
    End If
    panDetails.Enabled = True
    cmdOK1.Enabled = True
    cmdModify1.Enabled = False
    cmdCancel1.Enabled = True
    panCalcEnt.Visible = False
    cmdCalc.Caption = "Show Calculated Entitlements"
    dlpFDate1.Visible = True
    DateSeleV(1).Visible = False
    DateSeleV(2).Visible = False
    dlpFDate1.Top = DateSeleV(1).Top
    dlpTDate1.Visible = True
    dlpTDate1.Top = DateSeleV(2).Top
    dlpFDate1.SetFocus
    
    'Allow Sick Entitlement Date range to be modified.
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then   'Kawartha Lakes or Crown Investment Corp.(Ticket #14084)
        dlpFDate1S.Visible = True
        DateSeleS(1).Visible = False
        DateSeleS(2).Visible = False
        dlpFDate1S.Top = DateSeleS(1).Top
        dlpTDate1S.Visible = True
        dlpTDate1S.Top = DateSeleS(2).Top
    End If
End Sub

Private Sub cmdOK1_Click()
Dim xID, SQLQ
Dim rsTA As New ADODB.Recordset
    If Len(dlpFDate1) > 0 Then
       If Not IsDate(dlpFDate1) Then
           MsgBox "Not a valid date"
           dlpFDate1 = ""
           dlpFDate1.SetFocus
           Exit Sub
       End If
    Else
           MsgBox "Vacation From Date is required"
           dlpFDate1 = ""
           dlpFDate1.SetFocus
           Exit Sub
    End If
    If Len(dlpTDate1) > 0 Then
       If Not IsDate(dlpTDate1) Then
           MsgBox "Not a valid date"
           dlpTDate1 = ""
           dlpTDate1.SetFocus
           Exit Sub
       End If
    Else
           MsgBox "Vacation To Date is required"
           dlpTDate1 = ""
           dlpTDate1.SetFocus
           Exit Sub
    End If
    
    'Validate Sick Entitlement Date range
    'City of Kawartha Lakes or Crown Investment Corp.(Ticket #14084)
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then
        If Len(dlpFDate1S) > 0 Then
            If Not IsDate(dlpFDate1S) Then
                MsgBox "Not a valid date"
                dlpFDate1S = ""
                dlpFDate1S.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Sick Time From Date is required"
            dlpFDate1S = ""
            dlpFDate1S.SetFocus
            Exit Sub
        End If
        If Len(dlpTDate1S) > 0 Then
            If Not IsDate(dlpTDate1S) Then
                MsgBox "Not a valid date"
                dlpTDate1S = ""
                dlpTDate1S.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Sick Time To Date is required"
            dlpTDate1S = ""
            dlpTDate1S.SetFocus
            Exit Sub
        End If
    End If
    
    xID = Data1.Recordset("ED_EMPNBR")
    
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    
    'rsDATA.Resync
    'xID = rsDATA!ED_EMPNBR
        
    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & 0 & " WHERE ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
    'Save Sick Time Taken
    'City of Kawartha Lakes or Crown Investment Corp. (Ticket #14084)
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then
        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICKT=" & 0 & " WHERE ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
    End If

    SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
    SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3)='VAC' "
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpFDate1)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpTDate1)
    SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    Do Until rsTA.EOF
        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
        rsTA.MoveNext
    Loop
    rsTA.Close
       
    'Save Sick Time Taken
    'City of Kawartha Lakes or Crown Investment Corp (Ticket #14084)
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then
        SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
        SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3)='SIC' "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpFDate1S)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpTDate1S)
        SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
        SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
        rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HREMP SET ED_SICKT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
            rsTA.MoveNext
        Loop
        rsTA.Close
    End If
    
    'Call Set_Control("U", Me, rsDATA)
    
    Data1.Refresh
    Data1.Recordset.Find "ED_EMPNBR=" & xID
    
    panDetails.Enabled = False
    cmdOK1.Enabled = False
    cmdCancel1.Enabled = False
    cmdModify1.Enabled = True
    dlpFDate1.Visible = False
    dlpTDate1.Visible = False
    DateSeleV(1).Visible = True
    DateSeleV(2).Visible = True
    
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then  'City of Kawartha Lakes or Crown Investment Corp (Ticket #14084)
        dlpFDate1S.Visible = False
        dlpTDate1S.Visible = False
        DateSeleS(1).Visible = True
        DateSeleS(2).Visible = True
    End If

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport

'----------\\
    RHeading = Me.Caption
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading
    'Me.vbxCrystal(1).Action = 1
    If cmdDays.Enabled = False Then
        xReport = glbIHRREPORTS & "rgvacsic.rpt"
    Else
        xReport = glbIHRREPORTS & "rgvacsi1.rpt"
    End If
    
    Me.vbxCrystal.ReportFileName = xReport
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport

'----------\\
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    RHeading = Me.Caption
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading
    'Me.vbxCrystal(1).Action = 1
    If cmdDays.Enabled = False Then
        xReport = glbIHRREPORTS & "rgvacsic.rpt"
    Else
        xReport = glbIHRREPORTS & "rgvacsi1.rpt"
    End If
    
    Me.vbxCrystal.ReportFileName = xReport
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalc_Click(Index As Integer)
Dim Msg, Response, DgDef, SQLQ As String

Msg = "Do you wish to proceed and recalculate "
If Index = 1 Then
    Msg = Msg & "all Employees' "
Else
    Msg = Msg & "the Employee's "
End If
Msg = Msg & "outstanding entitlement ?"

'Ticket #20020
If glbEntOutStanding$ <> "1" Or glbEntOutStandingS$ <> "1" Then
    Msg = Msg & vbCrLf & vbCrLf & "NOTE: If the Entitlement Date Range of an employee has ended prior to today's date, "
    Msg = Msg & vbCrLf & "then the Entitlement Period for that employee will be change to new Entitlement Date Range."
End If

DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "ReCalculate")
If Response = IDNO Then Exit Sub

Screen.MousePointer = HOURGLASS
If Index = 1 Then
    Call EntReCalc("")
Else
    If glbGuelph Then   ' FOR Guelph-Willington
        Call AddFTE(Data1.Recordset("ED_EMPNBR"), "NEW")
    End If
    SQLQ = "ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    
    'County of Essex - Ticket #12676
    If glbCompSerial = "S/N - 2192W" Then
        Call EntReCalc(SQLQ, True)
    Else
        Call EntReCalc(SQLQ)
    End If
    If glbCompSerial = "S/N - 2173W" Then 'Town of Ajax 'Ticket #30402 Franks 08/02/2017
        Call Recalculate_OTBANK_Ajax_AllEmployees(Data1.Recordset("ED_EMPNBR"))
    End If
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Screen.MousePointer = DEFAULT

glbENTScreen = True
If Index = 1 Then
    Call Form_Activate
Else
    Data1.Recordset.Find SQLQ
End If

End Sub

Function EERetrieve()
Dim SQLQ As String, Q As QueryDef
'Dim db As Database
Dim countr   As Integer  ' EERetrieve_Snap is definded at form level

EERetrieve = False         ' if not found - no depts

SavEntOpt = glbEntOutStanding$

Select Case glbEntOutStanding$ ' sets field reference for basic 'which date'
    Case "2": fglbWDate$ = "ED_DOH"
    Case "3": fglbWDate$ = "ED_SENDTE"
    Case "4": fglbWDate$ = "ED_LTHIRE"
    Case "5": fglbWDate$ = "ED_USRDAT1"
    Case "6": fglbWDate$ = "ED_UNION"
End Select

SavEntOptS = glbEntOutStandingS$

Select Case glbEntOutStandingS$ ' sets field reference for basic 'which date'
    Case "2": fglbWDateS$ = "ED_DOH"
    Case "3": fglbWDateS$ = "ED_SENDTE"
    Case "4": fglbWDateS$ = "ED_LTHIRE"
    Case "5": fglbWDateS$ = "ED_USRDAT1"
    Case "6": fglbWDateS$ = "ED_UNION"
End Select

If glbEntOutStanding$ > "0" And glbEntOutStanding$ < "7" Then
    If glbEntOutStanding$ = "1" Then DateSeleV(0).Caption = "Entitlements Date"
    If glbEntOutStanding$ = "2" Then DateSeleV(0).Caption = lStr("Original Hire Date")
    If glbEntOutStanding$ = "3" Then DateSeleV(0).Caption = lStr("Seniority Date")
    If glbEntOutStanding$ = "4" Then DateSeleV(0).Caption = lStr("Last Hire Date")
    If glbEntOutStanding$ = "5" Then DateSeleV(0).Caption = lStr("User Defined Date")
    If glbEntOutStanding$ = "6" Then DateSeleV(0).Caption = lStr("Union Date")
End If

If glbEntOutStandingS$ > "0" And glbEntOutStandingS$ < "7" Then
    If glbEntOutStandingS$ = "1" Then DateSeleS(0).Caption = "Entitlements Date"
    If glbEntOutStandingS$ = "2" Then DateSeleS(0).Caption = lStr("Original Hire Date")
    If glbEntOutStandingS$ = "3" Then DateSeleS(0).Caption = lStr("Seniority Date")
    If glbEntOutStandingS$ = "4" Then DateSeleS(0).Caption = lStr("Last Hire Date")
    If glbEntOutStandingS$ = "5" Then DateSeleS(0).Caption = lStr("User Defined Date")
    If glbEntOutStandingS$ = "6" Then DateSeleS(0).Caption = lStr("Union Date")
End If

SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
If glbLinamar Then
    SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
Else
    If glbOracle Then
        SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
    End If
    
End If
SQLQ = SQLQ & "ED_EMPNBR,ED_VAC,ED_PVAC,"
SQLQ = SQLQ & "ED_SICK,ED_PSICK,ED_VACT,ED_SICKT,ED_ANNVAC, ED_ANNSICK, "
SQLQ = SQLQ & "ED_EFDATE,ED_EFDATES,ED_ETDATE,ED_ETDATES,"
SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,"
SQLQ = SQLQ & "ED_PT,ED_ORG,"
If glbLinamar Then
    SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
    SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
    SQLQ = SQLQ & " ED_VAC/8   AS WK_VACDAY, "
    SQLQ = SQLQ & " ED_PSICK/8 AS WK_PSICKDAY, "
    SQLQ = SQLQ & " ED_SICK/8  AS WK_SICKDAY, "
    SQLQ = SQLQ & " ED_VACT/8  AS WK_VACTDAY, "
    SQLQ = SQLQ & " ED_SICKT/8 AS WK_SICKTDAY, "
    SQLQ = SQLQ & " ED_ANNVAC/8   AS WK_CVDAY, "
    SQLQ = SQLQ & " ED_ANNSICK/8  AS WK_CSDAY, "
    SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
    SQLQ = SQLQ & "(ROUND([ED_VAC]/8,2)+ROUND([ED_PVAC]/8,2)-ROUND([ED_VACT]/8,2)) AS WK_VACODAY, "
    SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
    SQLQ = SQLQ & "(ROUND([ED_PSICK]/8,2)+ROUND([ED_SICK]/8,2)-ROUND([ED_SICKT]/8,2)) AS WK_SICKODAY "
ElseIf glbOracle Or glbSQL Then
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_VAC,2)+ROUND(ED_PVAC,2))-ROUND(ED_VACT,2)) + ROUND(ED_ANNVAC,2) END) AS WK_AVLVDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_PSICK,2)+ROUND(ED_SICK,2))-ROUND(ED_SICKT,2)) + ROUND(ED_ANNSICK,2) END) AS WK_AVLSDAY, "
        SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
        SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
    Else
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
        If Not glbCompSerial = "S/N - 2380W" Then 'VitalAire
            SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
            SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
        Else '#14635
            SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT-ED_OTBANK AS WK_VACO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2)-ROUND(ED_OTBANK/ED_DHRS,2) END) AS WK_VACODAY, "
            SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
        End If
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire
            SQLQ = SQLQ & ",ED_ANNVAC-ED_VACT AS WK_VACVLO "
            SQLQ = SQLQ & ",ED_ANNSICK-ED_SICKT AS WK_SICKVLO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((ED_ANNVAC-ED_VACT)/ED_DHRS,2) END) AS WK_VACVLODAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((ED_ANNSICK-ED_SICKT)/ED_DHRS,2) END) AS WK_SICKVLODAY "
        End If
    End If
Else
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PVAC]/[ED_DHRS]) AS WK_PVACDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VAC]/[ED_DHRS]) AS WK_VACDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PSICK]/[ED_DHRS]) AS WK_PSICKDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICK]/[ED_DHRS]) AS WK_SICKDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VACT]/[ED_DHRS]) AS WK_VACTDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICKT]/[ED_DHRS]) AS WK_SICKTDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNVAC]/[ED_DHRS]) AS WK_VCDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNSICK]/[ED_DHRS]) AS WK_VSDAY, "
    SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_VAC]+[ED_PVAC]-[ED_VACT])/[ED_DHRS]) AS WK_VACODAY, "
    SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_PSICK]+[ED_SICK]-[ED_SICKT])/[ED_DHRS]) AS WK_SICKODAY "
End If

If glbtermopen Then
    SQLQ = SQLQ & ",TERM_SEQ "
    SQLQ = SQLQ & " From Term_HREMP "
Else
    SQLQ = SQLQ & " From HREMP "
End If
SQLQ = SQLQ & "Where " & glbSeleDeptUn

If EESNameSort = True Then
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
Else
    SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
End If
    
Data1.RecordSource = SQLQ
Data1.Refresh
If glbtermopen Then
    If glbTERM_Seq > 0 Then
        SQLQ = "TERM_SEQ = " & glbTERM_Seq
        Data1.Recordset.Find SQLQ
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
            vbxTrueGridVLDay.DataSource = Data1
            vbxTrueGridVL.DataSource = Data1
        End If
    End If
Else
    If glbLEE_ID > 0 Then
        SQLQ = "ED_EMPNBR = " & glbLEE_ID
        Data1.Recordset.Find SQLQ
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
            vbxTrueGridVLDay.DataSource = Data1
            vbxTrueGridVL.DataSource = Data1
        End If
    End If
End If

EERetrieve = True
Exit Function

EERetrieve_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "VacList", "HREMP", "Select")
Call RollBack '28July99 js

End Function

Private Sub CmdRecalc_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdReCompDAccrual_Click()
    Dim Response%
    Dim bmk As Variant
    
    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If cmdReCompDAccrual.Visible = True Then
        If glbCompEntVacDaily Then
            'Comfirm the Re-Computation of Daily Accrual
            Response% = MsgBox("This function will create/recreate the Daily Accruals for this Employee as of Entitlement Start Date." & Chr(10) & Chr(10) & "Are you sure you want to proceed with this?", vbQuestion + vbYesNo, "Create Daily Accrual File")
            If Response% = IDNO Then
                Exit Sub
            End If
            
            Call Recompute_DailyAccrualFile(Data1.Recordset("ED_EMPNBR"), dlpFDate1.Text)
                        
            If Not IsDate(dlpFDate1.Text) Or Not IsDate(dlpTDate1.Text) Then
                MsgBox ("Failed to create Daily Accrual file for this Employee. Employee does not seem to belong to any Daily Accrual rule."), vbExclamation, "Failed to Create Daily Accrual"
            Else
                MsgBox ("Daily Accrual created for this employee successfully."), vbInformation, "Daily Accrual Created"
            End If
            
            Call Form_Activate
        End If
    End If
End Sub

Private Sub cmdReCompDAccrual_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpFDate1_Change()
DateSeleV(1) = dlpFDate1
End Sub

Private Sub dlpFDate1S_Change()
DateSeleS(1) = dlpFDate1S
End Sub

Private Sub dlpTDate1_Change()
DateSeleV(2) = dlpTDate1
End Sub

Private Sub dlpTDate1S_Change()
DateSeleS(2) = dlpTDate1S
End Sub

Private Sub dlpTmpEndDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpTmpEndDate_LostFocus()
    Dim xVacTaken
    Dim xDHRS
    
    If Len(dlpTmpEndDate) > 0 Then
        If Not IsDate(dlpTmpEndDate) Then
            MsgBox "Invalid End Date"
            dlpTmpEndDate.SetFocus
            Exit Sub
        Else
            'Get the new temp. Vacation TAKEN
            xVacTaken = Get_VacationTaken(Data1.Recordset("ED_EMPNBR"), DateSeleV(1), dlpTmpEndDate)
            
            'Display the new TAKEN
            If lblDayHrs.Caption = "DAYS" Then
                xDHRS = GetEmpData(Data1.Recordset("ED_EMPNBR"), "ED_DHRS")
                If xDHRS <> 0 And xDHRS <> "" And xVacTaken <> Empty Then
                    medTmpVTakenDy = xVacTaken / xDHRS
                Else
                    medTmpVTakenDy = ""
                End If
                medTmpVTakenHr = ""
                
                'Ticket #21665
                medVacR(1) = (IIf(IsNull(medPVac(1)), Val(0), Val(medPVac(1))) + IIf(IsNull(medCVacDay), Val(0), Val(medCVacDay))) - Val(medTmpVTakenDy)
            Else
                medTmpVTakenHr = IIf(IsNull(xVacTaken), 0, xVacTaken)
                medTmpVTakenDy = ""
                
                'Ticket #21665
                medVacR(0) = (IIf(IsNull(medPVac(0)), Val(0), Val(medPVac(0))) + IIf(IsNull(medCVac), Val(0), Val(medCVac))) - Val(medTmpVTakenHr)
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
Dim SQLQ

fbtnCalc = False
glbOnTop = "FRMVACSICKO"

If glbENTScreen = True Then
    glbENTScreen = False
    If EERetrieve() = False Then     ' get the info for this person
        Exit Sub
    End If          ' dpartment specific and populate the list
    If glbtermopen Then
        If glbTERM_Seq > 0 Then
            SQLQ = "TERM_SEQ = " & glbTERM_Seq
            Data1.Recordset.Find SQLQ
            If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
                vbxTrueGridVLDay.DataSource = Data1
                vbxTrueGridVL.DataSource = Data1
            End If
        End If
    Else
        If glbLEE_ID > 0 Then
            SQLQ = "ED_EMPNBR = " & glbLEE_ID
            Data1.Recordset.Find SQLQ
            If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
                vbxTrueGridVLDay.DataSource = Data1
                vbxTrueGridVL.DataSource = Data1
            End If
        End If
    End If
    
End If
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMVACSICKO"
End Sub

Private Sub Form_Load()
Dim SQLQ As String, EEID&, CompNo&
Dim x%

glbOnTop = "FRMVACSICKO"

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).Caption = "Retrieving Employee List - Stand by"

'7.9 Enhancement - only allow "Show Calculated" to be clicked if Monthly Entitlement Update
If glbCompEntVac$ = "M" Or glbCompEntSick$ = "M" Then
    cmdCalc.Enabled = True
Else
    cmdCalc.Enabled = False
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
    'Set ED_OTBANK to zero for the first time
    SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_OTBANK IS NULL"
    gdbAdoIhr001.Execute SQLQ
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    vbxTrueGrid.Visible = False
    vbxTrueGrid1.Visible = False
    vbxTrueGrid.DataSource = Data1
    vbxTrueGrid1.DataSource = Data1
    
    medVacRVL.DataField = "WK_VACVLO"
    medVacRVLDay.DataField = "WK_VACVLODAY"
    medSickR(3).DataField = "ED_SICK" '
    medSickR(2).DataField = "WK_SICKDAY"
    medCalcSick(3).DataField = "ED_SICKT"
    medCalcSick(2).DataField = "WK_SICKTDAY"
    medSickRVL.DataField = "WK_SICKO"
    medSickRVLDay.DataField = "WK_SICKODAY"
    
    'Relocate: panVADetails
    panVADetails.Top = panCalcEnt.Top
    panVADetails.Left = 90 '930
    panVADetails.Visible = True
    'Relocate: Calculated -> Sicktime
    medCalcSick(0).Top = 1100
    medCalcSick(1).Top = 1100
    'Relocate: Date Range -> Sicktime
    DateSeleS(1).Top = 1980
    DateSeleS(2).Top = 1980
    vbxTrueGridVL.Top = vbxTrueGrid.Top
    vbxTrueGridVL.Left = vbxTrueGrid.Left
    vbxTrueGridVL.Width = 10815
    vbxTrueGridVLDay.Top = vbxTrueGrid.Top
    vbxTrueGridVLDay.Left = vbxTrueGrid.Left
    vbxTrueGridVLDay.Width = 10815
    vbxTrueGridVL.Visible = True
    vbxTrueGridVL.Columns(3).DataField = "WK_VACO"
    vbxTrueGridVL.Columns(4).DataField = "ED_ANNVAC"
    vbxTrueGridVL.Columns(5).DataField = "WK_VACVLO"
    vbxTrueGridVL.Columns(6).DataField = "ED_SICK"
    vbxTrueGridVL.Columns(7).DataField = "ED_SICKT"
    vbxTrueGridVL.Columns(8).DataField = "WK_SICKO"
    vbxTrueGridVL.DataSource = Data1
    vbxTrueGridVLDay.Columns(3).DataField = "WK_VACODAY"
    vbxTrueGridVLDay.Columns(4).DataField = "WK_CVDAY"
    vbxTrueGridVLDay.Columns(5).DataField = "WK_VACVLODAY"
    vbxTrueGridVLDay.Columns(6).DataField = "WK_SICKDAY"
    vbxTrueGridVLDay.Columns(7).DataField = "WK_SICKTDAY"
    vbxTrueGridVLDay.Columns(8).DataField = "WK_SICKODAY"
    vbxTrueGridVLDay.DataSource = Data1

    vbxTrueGridVL.EditActive = False
    vbxTrueGridVL.MarqueeStyle = 4
    vbxTrueGridVLDay.EditActive = False
    vbxTrueGridVLDay.MarqueeStyle = 4
    
    cmdCalc.Visible = False  'VitalAire Ticket #13979
    
    If Not GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        lblTitle(11).Visible = False
        lblTitle(12).Visible = False
        medCalcVac(2).Enabled = False
        medCalcVac(3).Enabled = False
        panVCurrent.Enabled = False 'Ticket #14169
        medVacRVL.Visible = False
        medVacRVLDay.Visible = False
        medSickRVL.Visible = False
        medSickRVLDay.Visible = False
        vbxTrueGridVLDay.Columns(5).Visible = False
        vbxTrueGridVL.Columns(5).Visible = False
        vbxTrueGridVLDay.Columns(8).Visible = False
        vbxTrueGridVL.Columns(8).Visible = False
    Else
        lblTitle(11).Visible = True
        lblTitle(12).Visible = True
        medCalcVac(2).Enabled = True
        medCalcVac(3).Enabled = True
        panVCurrent.Enabled = True 'Ticket #14169
        medVacRVL.Visible = True
        medVacRVLDay.Visible = False
        medSickRVL.Visible = True
        medSickRVLDay.Visible = False
        vbxTrueGridVLDay.Columns(5).Visible = True
        vbxTrueGridVL.Columns(5).Visible = True
        vbxTrueGridVLDay.Columns(8).Visible = True
        vbxTrueGridVL.Columns(8).Visible = True
    End If
End If

If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
    cmdCalc.Enabled = True
    cmdCalc.Caption = "Show To Be Accrued"
    'lblTitle(6).Caption = "To be Accrued"
    'lblTitle(18).Caption = "Available"
    lblTitle(6).Visible = False
    lblTitle(18).Visible = False
    lblTitle(26).Visible = False
    lblTitle(27).Visible = False
    lblTitle(23).Caption = "To be Accrued"
    lblTitle(24).Caption = "Available"
    
    'vbxTrueGrid.Columns(7).Caption = "Vac. Avail."
    'vbxTrueGrid.Columns(12).Caption = "Sick Avail."
    'vbxTrueGrid1.Columns(7).Caption = "Vac. Avail."
    'vbxTrueGrid1.Columns(12).Caption = "Sick Avail."
    'vbxTrueGrid.Columns(7).DataField = "WK_CVDAY"
    'vbxTrueGrid.Columns(12).DataField = "WK_CSDAY"
    'vbxTrueGrid1.Columns(7).DataField = "WK_CVDAY"
    'vbxTrueGrid1.Columns(12).DataField = "WK_CSDAY"
    vbxTrueGrid.Columns(7).Visible = False      '"WK_CVDAY"
    vbxTrueGrid.Columns(13).Visible = False     '"WK_CSDAY"
    vbxTrueGrid1.Columns(7).Visible = False     '"WK_CVDAY"
    vbxTrueGrid1.Columns(13).Visible = False    '"WK_CSDAY"
    
    vbxTrueGrid.Columns(8).Visible = True     '"WK_AVLVDAY"
    vbxTrueGrid.Columns(14).Visible = True    '"WK_AVLSDAY"
    vbxTrueGrid1.Columns(8).Visible = True      '"WK_AVLVDAY"
    vbxTrueGrid1.Columns(14).Visible = True     '"WK_AVLSDAY"
Else
    vbxTrueGrid.Columns(8).Visible = False      '"WK_AVLVDAY"
    vbxTrueGrid.Columns(14).Visible = False     '"WK_AVLSDAY"
    vbxTrueGrid1.Columns(8).Visible = False      '"WK_AVLVDAY"
    vbxTrueGrid1.Columns(14).Visible = False     '"WK_AVLSDAY"
End If

'Ticket #23878 - Carizon Family and Community Services/KidsLink
If glbCompSerial = "S/N - 2430W" Then
    'lblTitle(6).Caption = "Annual Entitl."
    lblTitle(6).Visible = False
    lblTitle(26).Visible = False
    lblTitle(23).Caption = "Annual Entitl."
    cmdCalc.Caption = "Show Annual Entitlements"
End If

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

EESNameSort = True  'first sort is by surname
glbENTScreen = True     'Refresh DATA1 ... FORM.ACTIVATE

If glbCompSerial = "S/N - 2235W" Then   'laura 03/05/98
    lblDayHrs.Caption = "DAYS"
    cmdDays.Visible = False
    cmdHours.Visible = False
ElseIf glbCompSerial = "S/N - 2236W" Then
    lblDayHrs.Caption = "DAYS"
    cmdDays.Visible = False
    cmdHours.Visible = False
Else
    lblDayHrs.Caption = "HOURS"
End If

If glbCompSerial = "S/N - 2262W" Then
    medPVac(0).Format = "#,##0.0000"
    medPVac(1).Format = "#,##0.0000"
    medCVacDay.Format = "#,##0.0000"
    medCVac.Format = "#,##0.0000"
    medVacC(0).Format = "#,##0.0000"
    medVacC(1).Format = "#,##0.0000"
    medVacR(0).Format = "#,##0.0000"
    medVacR(1).Format = "#,##0.0000"
    medCalcVac(0).Format = "#,##0.0000"
    medCalcVac(1).Format = "#,##0.0000"
    medAvailVac(0).Format = "#,##0.0000"
    medAvailVac(1).Format = "#,##0.0000"
    vbxTrueGrid.Columns(3).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(4).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(5).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(6).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(12).NumberFormat = "0.0000"
    vbxTrueGrid1.Columns(3).NumberFormat = "0.0000"
    vbxTrueGrid1.Columns(4).NumberFormat = "0.0000"
    vbxTrueGrid1.Columns(5).NumberFormat = "0.0000"
    vbxTrueGrid1.Columns(6).NumberFormat = "0.0000"
    vbxTrueGrid1.Columns(12).NumberFormat = "0.0000"
End If

'For Essex Library and Kawartha Lakes and Crown Investment Corp (Ticket #14084)
If glbCompSerial = "S/N - 2296W" Or glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2205W" Then
    cmdModify1.Visible = True
    cmdOK1.Visible = True
    cmdCancel1.Visible = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Me.Show
Screen.MousePointer = DEFAULT

If Not gSec_Upd_Entitlements Then  'js
    CmdRecalc(0).Enabled = False
    CmdRecalc(1).Enabled = False
End If
If glbtermopen Then
    CmdRecalc(0).Visible = False
    CmdRecalc(1).Visible = False
End If

'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
If Not glbtermopen Then
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Visible = True
    Else
        cmdReCompDAccrual.Visible = False
    End If
End If

If glbWFC Or glbCompSerial = "S/N - 2418W" Then
    lblTitle(5).Visible = False
    DateSeleS(0).Visible = False
    DateSeleS(1).Visible = False
    DateSeleS(2).Visible = False
    medCSick.Visible = False
    medCSickDay.Visible = False
    medPSick(0).Visible = False
    medPSick(1).Visible = False
    medSickC(0).Visible = False
    medSickC(1).Visible = False
    medSickR(0).Visible = False
    medSickR(1).Visible = False
    medCalcSick(0).Visible = False
    medCalcSick(1).Visible = False
    
    medAvailSick(0).Visible = False
    medAvailSick(1).Visible = False
    
    vbxTrueGrid.Columns(8).Visible = False
    vbxTrueGrid.Columns(9).Visible = False
    vbxTrueGrid.Columns(10).Visible = False
    vbxTrueGrid.Columns(11).Visible = False
    vbxTrueGrid.Columns(12).Visible = False
    vbxTrueGrid1.Columns(8).Visible = False
    vbxTrueGrid1.Columns(9).Visible = False
    vbxTrueGrid1.Columns(10).Visible = False
    vbxTrueGrid1.Columns(11).Visible = False
    vbxTrueGrid1.Columns(12).Visible = False
    
    vbxTrueGrid.Columns(13).Visible = False     '"WK_AVLSDAY"
    vbxTrueGrid1.Columns(13).Visible = False     '"WK_AVLSDAY"
    
    Me.Caption = "Vacation Overview"
    
End If

End Sub

Private Sub medCalcVac_Change(Index As Integer)
    If medCalcVac(Index).Visible = True Or medCalcSick(Index).Visible = True Then
        If Index = 0 Then
            'Vacation
            If medVacR(1) = "" Or medCVacDay = "" Or medCalcVac(0) = "" Then
                medAvailVac(0).Text = ""
            Else
                'Ticket #23878 - KidsLink/Carizon - their Available will be Outstanding.
                If glbCompSerial = "S/N - 2430W" Then
                    'Available Days = Outstanding
                    medAvailVac(0).Text = medVacR(1)
                Else
                    'Available Days = (Outstanding - Current) + Calculated
                    medAvailVac(0).Text = (medVacR(1) - medCVacDay) + medCalcVac(0)
                End If
            End If
            
            'Sick
            If medSickR(1) = "" Or medCSickDay = "" Or medCalcSick(0) = "" Then
                medAvailSick(0).Text = ""
            Else
                'Ticket #23878 - KidsLink/Carizon - their Available will be Outstanding.
                If glbCompSerial = "S/N - 2430W" Then
                    'Available Days = Outstanding
                    medAvailSick(0).Text = medSickR(1)
                Else
                    'Available Days = (Outstanding - Current) + Calculated
                    medAvailSick(0).Text = (medSickR(1) - medCSickDay) + medCalcSick(0)
                End If
            End If
        ElseIf Index = 1 Then
            'Vacation
            If medVacR(0) = "" Or medCVac = "" Or medCalcVac(1) = "" Then
                medAvailVac(1).Text = ""
            Else
                'Ticket #23878 - KidsLink/Carizon - their Available will be Outstanding.
                If glbCompSerial = "S/N - 2430W" Then
                    'Available Days = Outstanding
                    medAvailVac(1).Text = medVacR(0)
                Else
                    'Available Days = (Outstanding - Current) + Calculated
                    medAvailVac(1).Text = (medVacR(0) - medCVac) + medCalcVac(1)
                End If
            End If
            
            'Sick
            If medSickR(0) = "" Or medCSick = "" Or medCalcSick(1) = "" Then
                medAvailSick(1).Text = ""
            Else
                'Ticket #23878 - KidsLink/Carizon - their Available will be Outstanding.
                If glbCompSerial = "S/N - 2430W" Then
                    'Available Days = Outstanding
                    medAvailSick(1).Text = medSickR(0)
                Else
                    'Available Days = (Outstanding - Current) + Calculated
                    medAvailSick(1).Text = (medSickR(0) - medCSick) + medCalcSick(1)
                End If
            End If
        End If
    ElseIf glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        If (medAvailVac(0).Visible = True Or medAvailVac(1).Visible = True) Then
            If Index = 0 Then
                'Days
                If medVacR(1) = "" Or medToBeVac(0) = "" Then
                    medAvailVac(0).Text = ""
                Else
                    'medAvailVac(0).Text = (medVacR(1) - medCVacDay) + medCalcVac(0)
                    medAvailVac(0).Text = Val(medVacR(1)) + Val(medToBeVac(0))
                End If
                
                If medSickR(1) = "" Or medToBeSick(0) = "" Then
                    medAvailSick(0).Text = ""
                Else
                    'medAvailSick(0).Text = (medSickR(1) - medCSickDay) + medCalcSick(0)
                    medAvailSick(0).Text = Val(medSickR(1)) + Val(medToBeSick(0))
                End If
            ElseIf Index = 1 Then
                'Hours
                If medVacR(0) = "" Or medToBeVac(1) = "" Then
                    medAvailVac(1).Text = ""
                Else
                    'medAvailVac(1).Text = (medVacR(0) - medCVac) + medCalcVac(1)
                    medAvailVac(1).Text = Val(medVacR(0)) + Val(medToBeVac(1))
                End If
                
                If medSickR(0) = "" Or medToBeSick(1) = "" Then
                    medAvailSick(1).Text = ""
                Else
                    'medAvailSick(1).Text = (medSickR(0) - medCSick) + medCalcSick(1)
                    medAvailSick(1).Text = Val(medSickR(0)) + Val(medToBeSick(1))
                End If
            End If
        End If
    End If
End Sub

Private Sub medCSick_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCVac_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPSick_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPVac_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEESearch_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

    glbLEE_ID = Data1.Recordset("ED_EMPNBR")
    glbLEE_FName = Data1.Recordset("ED_FNAME")
    glbLEE_SName = Data1.Recordset("ED_SURNAME")
    If glbLinamar Then
        glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
    End If
End Sub

Private Sub vbxTrueGridVL_DblClick()
    glbLEE_ID = Data1.Recordset("ED_EMPNBR")
    glbLEE_FName = Data1.Recordset("ED_FNAME")
    glbLEE_SName = Data1.Recordset("ED_SURNAME")
    If glbLinamar Then
        glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
    End If
End Sub

Private Sub vbxTrueGridVLDay_DblClick()
    glbLEE_ID = Data1.Recordset("ED_EMPNBR")
    glbLEE_FName = Data1.Recordset("ED_FNAME")
    glbLEE_SName = Data1.Recordset("ED_SURNAME")
    If glbLinamar Then
        glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
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


Sub AddFTE(xEmpNo, xFLAG)
    Dim OldFTE, NewFTE, xEFDATE, xETDATE, xNumVac
    Dim fNewFTE, fOldFTE, FlagOldFTE
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
    If RsFTEHis.EOF And xFLAG <> "NEW" Then
        Exit Sub
    End If

    If xFLAG = "NEW" Then
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
        If IsNull(RsJobEmp("JH_FTENUM")) Then
            fNewFTE = 0
        Else
            fNewFTE = RsJobEmp("JH_FTENUM")
        End If
        RsJobEmp.MoveNext
        fOldFTE = 0
        FlagOldFTE = True
        II = 0
        Do While (Not RsJobEmp.EOF) And FlagLoop
            xDate1 = RsJobEmp("JH_SDATE")
            If FlagOldFTE Then
                If Not IsNull(RsJobEmp("JH_FTENUM")) Then
                    fOldFTE = RsJobEmp("JH_FTENUM")
                End If
                FlagOldFTE = False
            End If
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
    
    xVacDays = 0
    For J = 1 To II
        xVacDays = xVacDays + xArray(J, 2)
    Next
    
    If xVacDays = 0 Then
        Exit Sub
    End If
    xVacHours = Round(xVacDays, 0) * xHrsDay
    
    If xVacHours <> xNumVacINS Then
        gdbAdoIhr001.BeginTrans
        'Dim RsTempEmp As New ADODB.Recordset
        SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
        RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        If Not RsTempEmp.EOF Then
            RsTempEmp("ED_VAC") = xVacHours
            RsTempEmp.Update
        End If
        RsTempEmp.Close
        gdbAdoIhr001.CommitTrans
        
        If RsFTEHis.EOF Then
            RsFTEHis.AddNew
            RsFTEHis("CP_EMPNBR") = xEmpNo
            RsFTEHis("CP_VACORIGION") = xNumVac
            RsFTEHis("CP_VACO") = xNumVacINS
            RsFTEHis("CP_VACN") = xVacHours
            If fOldFTE > 0 Then
            RsFTEHis("CP_FTENUMO") = fOldFTE
            End If
            If fNewFTE > 0 Then
            RsFTEHis("CP_FTENUMN") = fNewFTE
            End If
            RsFTEHis("CP_FDATE") = CVDate(xEFDATE)
            RsFTEHis("CP_TDATE") = CVDate(xETDATE)
            RsFTEHis("CP_LDATE") = Date
            RsFTEHis("CP_LTIME") = Time$
            RsFTEHis("CP_LUSER") = glbUserID
            RsFTEHis.Update
        End If
    End If
    RsFTEHis.Close
    
    Exit Sub

ExitLin1:
End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ

fbtnCalc = True

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
    If glbCompEntVacDaily Then
        cmdReCompDAccrual.Enabled = False
    End If
    
    Exit Sub
End If
    
'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
If glbCompEntVacDaily Then
    cmdReCompDAccrual.Enabled = True
End If
    
SQLQ = Data1.RecordSource
SQLQ = Left(SQLQ, InStr(SQLQ, "ORDER BY") - 1)
SQLQ = SQLQ & " AND ED_EMPNBR= " & Data1.Recordset!ED_EMPNBR
If glbtermopen Then
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)

End Sub


Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
        If glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            If glbOracle Then
                SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
            Else
                SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            End If
            
        End If
        SQLQ = SQLQ & "ED_EMPNBR,ED_VAC,ED_PVAC,"
        SQLQ = SQLQ & "ED_SICK,ED_PSICK,ED_VACT,ED_SICKT,ED_ANNVAC, ED_ANNSICK, "
        SQLQ = SQLQ & "ED_EFDATE,ED_EFDATES,ED_ETDATE,ED_ETDATES,"
        SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,"
        SQLQ = SQLQ & "ED_PT,ED_ORG,"
        If glbLinamar Then
            SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
            SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
            SQLQ = SQLQ & " ED_VAC/8   AS WK_VACDAY, "
            SQLQ = SQLQ & " ED_PSICK/8 AS WK_PSICKDAY, "
            SQLQ = SQLQ & " ED_SICK/8  AS WK_SICKDAY, "
            SQLQ = SQLQ & " ED_VACT/8  AS WK_VACTDAY, "
            SQLQ = SQLQ & " ED_SICKT/8 AS WK_SICKTDAY, "
            SQLQ = SQLQ & " ED_ANNVAC/8   AS WK_CVDAY, "
            SQLQ = SQLQ & " ED_ANNSICK/8  AS WK_CSDAY, "
            SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
            SQLQ = SQLQ & "(ROUND([ED_VAC]/8,2)+ROUND([ED_PVAC]/8,2)-ROUND([ED_VACT]/8,2)) AS WK_VACODAY, "
            SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
            SQLQ = SQLQ & "(ROUND([ED_PSICK]/8,2)+ROUND([ED_SICK]/8,2)-ROUND([ED_SICKT]/8,2)) AS WK_SICKODAY "
        ElseIf glbOracle Or glbSQL Then
            If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2))-ROUND(ED_VACT/ED_DHRS,2)) + ROUND(ED_ANNVAC/ED_DHRS,2) END) AS WK_AVLVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2))-ROUND(ED_SICKT/ED_DHRS,2)) + ROUND(ED_ANNSICK/ED_DHRS,2) END) AS WK_AVLSDAY, "
                SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
                SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
            Else
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
                SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
                SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
            End If
        Else
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PVAC]/[ED_DHRS]) AS WK_PVACDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VAC]/[ED_DHRS]) AS WK_VACDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PSICK]/[ED_DHRS]) AS WK_PSICKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICK]/[ED_DHRS]) AS WK_SICKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VACT]/[ED_DHRS]) AS WK_VACTDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICKT]/[ED_DHRS]) AS WK_SICKTDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNVAC]/[ED_DHRS]) AS WK_VCDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNSICK]/[ED_DHRS]) AS WK_VSDAY, "
            SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_VAC]+[ED_PVAC]-[ED_VACT])/[ED_DHRS]) AS WK_VACODAY, "
            SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_PSICK]+[ED_SICK]-[ED_SICKT])/[ED_DHRS]) AS WK_SICKODAY "
        End If
        
        If glbtermopen Then
            SQLQ = SQLQ & ",TERM_SEQ "
            SQLQ = SQLQ & " From Term_HREMP "
        Else
            SQLQ = SQLQ & " From HREMP "
        End If
        SQLQ = SQLQ & "Where " & glbSeleDeptUn
       
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    chkEndDateEdit.Value = False
    
    If fbtnCalc = False And cmdCalc.Caption = "Show Entitlement Dates" Then
        Call ReCalcAnn
        DoEvents
    End If
    Call Display_Value
    fbtnCalc = False
    
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
    End If

End Sub

Private Sub vbxTrueGrid1_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid1.Tag = "ASC" Then
            vbxTrueGrid1.Tag = "DESC"
        Else
            vbxTrueGrid1.Tag = "ASC"
        End If
        
        SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
        If glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            If glbOracle Then
                SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
            Else
                SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            End If
            
        End If
        SQLQ = SQLQ & "ED_EMPNBR,ED_VAC,ED_PVAC,"
        SQLQ = SQLQ & "ED_SICK,ED_PSICK,ED_VACT,ED_SICKT,ED_ANNVAC, ED_ANNSICK, "
        SQLQ = SQLQ & "ED_EFDATE,ED_EFDATES,ED_ETDATE,ED_ETDATES,"
        SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,"
        SQLQ = SQLQ & "ED_PT,ED_ORG,"
        If glbLinamar Then
            SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
            SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
            SQLQ = SQLQ & " ED_VAC/8   AS WK_VACDAY, "
            SQLQ = SQLQ & " ED_PSICK/8 AS WK_PSICKDAY, "
            SQLQ = SQLQ & " ED_SICK/8  AS WK_SICKDAY, "
            SQLQ = SQLQ & " ED_VACT/8  AS WK_VACTDAY, "
            SQLQ = SQLQ & " ED_SICKT/8 AS WK_SICKTDAY, "
            SQLQ = SQLQ & " ED_ANNVAC/8   AS WK_CVDAY, "
            SQLQ = SQLQ & " ED_ANNSICK/8  AS WK_CSDAY, "
            SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
            SQLQ = SQLQ & "(ROUND([ED_VAC]/8,2)+ROUND([ED_PVAC]/8,2)-ROUND([ED_VACT]/8,2)) AS WK_VACODAY, "
            SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
            SQLQ = SQLQ & "(ROUND([ED_PSICK]/8,2)+ROUND([ED_SICK]/8,2)-ROUND([ED_SICKT]/8,2)) AS WK_SICKODAY "
        ElseIf glbOracle Or glbSQL Then
            If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_VAC,2)+ROUND(ED_PVAC,2))-ROUND(ED_VACT,2)) + ROUND(ED_ANNVAC,2) END) AS WK_AVLVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_PSICK,2)+ROUND(ED_SICK,2))-ROUND(ED_SICKT,2)) + ROUND(ED_ANNSICK,2) END) AS WK_AVLSDAY, "
                SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
                SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
            Else
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
                SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
                SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
            End If
        Else
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PVAC]/[ED_DHRS]) AS WK_PVACDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VAC]/[ED_DHRS]) AS WK_VACDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PSICK]/[ED_DHRS]) AS WK_PSICKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICK]/[ED_DHRS]) AS WK_SICKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VACT]/[ED_DHRS]) AS WK_VACTDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICKT]/[ED_DHRS]) AS WK_SICKTDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNVAC]/[ED_DHRS]) AS WK_VCDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNSICK]/[ED_DHRS]) AS WK_VSDAY, "
            SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_VAC]+[ED_PVAC]-[ED_VACT])/[ED_DHRS]) AS WK_VACODAY, "
            SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_PSICK]+[ED_SICK]-[ED_SICKT])/[ED_DHRS]) AS WK_SICKODAY "
        End If
        
        If glbtermopen Then
            SQLQ = SQLQ & ",TERM_SEQ "
            SQLQ = SQLQ & " From Term_HREMP "
        Else
            SQLQ = SQLQ & " From HREMP "
        End If
        SQLQ = SQLQ & "Where " & glbSeleDeptUn
       
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid1.Columns(ColIndex).DataField & " " & vbxTrueGrid1.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    chkEndDateEdit.Value = False

    If fbtnCalc = False And cmdCalc.Caption = "Show Entitlement Dates" Then
        Call ReCalcAnn
        DoEvents
    End If
    Call Display_Value
    fbtnCalc = False

    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
    End If
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
UpdateRight = GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
    Updateble = GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID)
Else
    Updateble = False
End If
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    UpdateState = OPENING
    TF = True
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Private Function modAnnVacation(isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012

On Error GoTo modUpdateSelection_Err

modAnnVacation = False

'gdbAdoIhr001.BeginTrans


    if_Entitle = False
    if_Vacation = False

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

    If snapEntitle.State = 0 Then Exit Function
    
    empNo& = snapEntitle("ED_EMPNBR")

    If IsNull(snapEntitle("ED_ANNVAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_ANNVAC")
    End If

    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If

    spt = snapEntitle("ED_PT")

    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)

    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    'rsJOB.Close    'Ticket #22842 -moved below because of calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    If glbLinamar Then dblDHours# = 8

    xAsOf = fglbAsOf
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1

    If rsEntRules("VE_EMONTH") > 0 Then
        If dblServiceYears# >= CDbl(rsEntRules("VE_BMONTH")) And dblServiceYears# <= CDbl(rsEntRules("VE_EMONTH")) Then
            intWhereFit& = x%
            If Len(rsEntRules("VE_ENTITLE")) > 0 Then if_Entitle = True
            If Len(rsEntRules("VE_PCT")) > 0 Then if_Vacation = True
        End If
    End If

    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges

    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    'Ticket #22842 - calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If rsEntRules("VE_TYPE") = "D" Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If rsEntRules("VE_TYPE") = "F" Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close

    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
    ' which represents if Sick and Vacation entitlements
    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
    ' and read on system startup.

    ' In this routine we work independantly of SICK/VACATIon entitlement.
    '  fglbCompMonthly% - is the independant representation
        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
        'Procedure modUpdateSelection is used to set
        'fglbCompMonthly based on values it finds for global variables
        ' and what the user wants to manipulate (sick/Vac)

    'optD indicates if Entitlement entered is Daily or yearly based
    ' if daily then max entitlement is based on entitlement * hours they work.

    ' we have   Entitle = existing entitmenet (stored presently
    '           NewEntitle = amount entered onto screen = medentitle(index)
    '           EntitleUpd  = value to update record with

    If if_Entitle Then
        dblNewEntitle# = rsEntRules("VE_ENTITLE")
        dblNewMax# = 0
        If rsEntRules("VE_TYPE") = "D" Then           ' Entitlements entered in days
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
                If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "F" Then
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
                If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "H" Then
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX")
        End If
        dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values

         If dblNewMax <> 0 Then          'only do if not zero
            'Ticket #23878 - KidsLink/Carizon - their Calculated will be Annualized Vacation not using Prev.
            If glbCompSerial = "S/N - 2430W" Then
                If dblEntitleUpd > dblNewMax Then
                    dblEntitleUpd = dblNewMax
                End If
            Else
                If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                    dblEntitleUpd = dblNewMax - dblPrevEntitle#
                End If
            End If
        End If

        DtTm = Now
    End If

    If if_Vacation Then
        If glbCBrant And Len(rsEntRules("VE_SECTION")) > 0 And snapEntitle("ED_SECTION") >= rsEntRules("VE_SECTION") Then
            VacpcN = rsEntRules("VE_PCT") + dblEntitle#
        Else
            VacpcN = rsEntRules("VE_PCT")
        End If
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(rsEntRules("VE_PCT")) Then snapEntitle("ED_VACPC") = rsEntRules("VE_PCT")

    End If
    If if_Entitle Then

        'If glbCompSerial = "S/N - 2188W" Then  'Ticket #8887
        '    dblEntitleUpd = Round(dblEntitleUpd, 0)
        If glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
        If glbCBrant And Len(rsEntRules("VE_SECTION")) > 0 Then
            dblEntitleUpd = rsEntRules("VE_PCT") + dblEntitle#
        End If


        If isLast And glbCompSerial = "S/N - 2376W" Then '#9536 on Oct 21,2005 George
            If dblDHours# <> 0 Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round((dblEntitleDays / 0.25 + 0.1), 0) * 0.25 ' round to 1/4 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            End If
        End If

        'Hemu - 12/31/2003 End
        'Added by bryan 13/Jun/06 Ticket#10916
        If dblEntitle# = 0 And glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
            snapEntitle("ED_ANNVAC") = IIf(dblEntitleUpd < 0, 0, dblEntitleUpd) '+ IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))
        Else
            If (IsNull(snapEntitle("ED_ANNVAC")) Or snapEntitle("ED_ANNVAC") = 0) Or (UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N") Then
                'Do not calculate Annual Vacation if it already has value. This is because if Rounded at the time of Vacation Ent. Upd
                'then in this routine it will not do that. Also Ann Vac is calculated at the time of Vac Update.
                snapEntitle("ED_ANNVAC") = dblEntitleUpd
            End If
        End If
    End If
    snapEntitle.Update
    

lblNextRec:
'gdbAdoIhr001.CommitTrans
modAnnVacation = True

'Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub ReCalcAnn()
    'Rules will be looked at in order of creation. If a rule affects everyone AFTER
    'a rule affects a group it could cause problems with the calculated annual value

    Dim SQLQ As String
    Dim prec As Long, pct As Long, lngRecs As Long, xRunTimes As Integer
    Dim blIsLast As Boolean
    Dim bmk As Variant
    Dim xBalMonths As Integer
    Dim fglbLastAccDt As Variant
    Dim fglbEntitlEndDt As Date
    Dim flgAllowAnnual As Boolean
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #29230 - Daily Entitlement - already calculated and coming from different table
    If glbCompEntVac$ = "D" Then
        'Do not Recalculate for Vacation for Daily Entitlement as it's already calculated from different entitlement table
        'So get the Annual Accrual from the Daily Accrual table
        Call Update_AnnualVac_From_DailyAccrual
        
        'ReCalculate and Update Annual Sick Time
        GoTo SickTime
    End If
    
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0"
    'Else   'Don't want to reset to 0 because the Annual Vac has been calculated from Vac Entitl Master screen and may have Rounded option applied.
    '    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0 WHERE ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
    End If
    
    flgAllowAnnual = False
    
    SQLQ = "SELECT * FROM HRVACENT ORDER BY VE_ID"
    Set rsEntRules = Nothing
    rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsEntRules.EOF Then
        rsEntRules.MoveFirst
        
        'For each entitlement rule
        Do While Not rsEntRules.EOF
            If (UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N") And Not (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0) Then
                'Monthly or Annualized Monthly
                prec = 0
                If Not CR_SnapEntitle("VAC") Then Exit Sub  ' create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    xBalMonths = 12
                    
                    'For the employee
                    Do While Not snapEntitle.EOF
                        If IsNull(snapEntitle("ED_EFDATE")) = False Then
                            'fglbAsOf = snapEntitle("ED_EFDATE")
                            fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATE")) & "/" & Day(rsEntRules("VE_EDATE")) & "/" & Year(snapEntitle("ED_EFDATE")), "mm/dd/yyyy"), Day(rsEntRules("VE_EDATE")), month(snapEntitle("ED_EFDATE")), Year(snapEntitle("ED_EFDATE")))
                            
                            If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
                                'Get last Accrual date from HR_ACCRUAL, and entitlement end date
                                fglbLastAccDt = Get_Employee_Last_Accrual_Date(snapEntitle("ED_EMPNBR"), "VAC")
                                fglbEntitlEndDt = snapEntitle("ED_ETDATE")
                                
                                'Calculate entitlement date - following month after the last update date found
                                'in the HR_ACCRUAL
                                If fglbLastAccDt = "" Then
                                    fglbAsOf = snapEntitle("ED_EFDATE")
                                Else
                                    fglbAsOf = DateAdd("m", 1, CVDate(fglbLastAccDt))
                                End If
                                
                                'Get the # of months left to end of entitlement so that the entitlement calc.
                                'is run that many times
                                xBalMonths = DateDiff("m", DateAdd("m", -1, fglbAsOf), fglbEntitlEndDt)
                            End If
                        Else
                            GoTo Fvac_NextEmp
                        End If
                        
                        If (glbCompSerial = "S/N - 2395W") Or (IsNull(snapEntitle("ED_ANNVAC")) Or snapEntitle("ED_ANNVAC") = 0) Or (flgAllowAnnual = True And Not IsNull(snapEntitle("ED_ANNVAC")) And snapEntitle("ED_ANNVAC") <> 0) Then
                        'Do not calculate Annual Vacation if it already has value. This is because if Rounded at the time of Vacation Ent. Upd
                        'then in this routine it will not do that. Also Ann Vac is calculated at the time of Vac Update.
                            For xRunTimes = 1 To xBalMonths
                                flgAllowAnnual = True
                                blIsLast = False
                                
                                If xRunTimes = xBalMonths Then blIsLast = True
                                
                                If Not modAnnVacation(blIsLast) Then GoTo Fvac
    
                                fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                            Next
                        End If
Fvac_NextEmp:
                        snapEntitle.MoveNext
                        
                    Loop    'Wend
                End If
                snapEntitle.Close
            Else
                'Annual Entitlement
                prec = 0
                If Not CR_SnapEntitle("VAC") Then Exit Sub  ' create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount

                        If IsNull(snapEntitle("ED_EFDATE")) = False Then
                            fglbAsOf = snapEntitle("ED_EFDATE")
                        Else
                            GoTo Fvac_NextEmp_1
                        End If

                        If Not modAnnVacation(True) Then GoTo Fvac
Fvac_NextEmp_1:
                        snapEntitle.MoveNext
                    Wend

                End If
                snapEntitle.Close
            End If
            
            rsEntRules.MoveNext
            
        Loop 'Until rsEntRules.EOF
        flgAllowAnnual = False
    End If
    
Fvac:
    rsEntRules.Close
    Set rsEntRules = Nothing
    
SickTime:
    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0"
    'Else   'Don't want to reset to 0 because the Annual Sick has been calculated from Sick Entitl Master screen and may have Rounded option applied.
    '    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
    End If
    
    flgAllowAnnual = False
    
    SQLQ = "SELECT * FROM HRSICKENT ORDER BY VE_ID"
    rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsEntRules.EOF Then
        rsEntRules.MoveFirst
        
        Do While Not rsEntRules.EOF
            If (UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N") And Not (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0) Then
                'Monthly or Annualized Monthly
                prec = 0
                If Not CR_SnapEntitle("SICK") Then Exit Sub  ' create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    xBalMonths = 12
                    
                    Do While Not snapEntitle.EOF
                        If IsNull(snapEntitle("ED_EFDATES")) = False Then
                            'fglbAsOf = snapEntitle("ED_EFDATES")
                            fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATES")) & "/" & Day(rsEntRules("VE_EDATE")) & "/" & Year(snapEntitle("ED_EFDATES")), "mm/dd/yyyy"), Day(rsEntRules("VE_EDATE")), month(snapEntitle("ED_EFDATES")), Year(snapEntitle("ED_EFDATES")))
                            
                            If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
                                'Get last Accrual date from HR_ACCRUAL, and entitlement end date
                                fglbLastAccDt = Get_Employee_Last_Accrual_Date(snapEntitle("ED_EMPNBR"), "SICK")
                                fglbEntitlEndDt = snapEntitle("ED_ETDATES")
                                
                                'Calculate entitlement date - following month after the last update date found
                                'in the HR_ACCRUAL
                                If fglbLastAccDt = "" Then
                                    fglbAsOf = snapEntitle("ED_EFDATES")
                                Else
                                    fglbAsOf = DateAdd("m", 1, CVDate(fglbLastAccDt))
                                End If
                                
                                'Get the # of months left to end of entitlement so that the entitlement calc.
                                'is run that many times
                                xBalMonths = DateDiff("m", DateAdd("m", -1, fglbAsOf), fglbEntitlEndDt)
                            End If
                        Else
                            GoTo Final
                        End If
                                                                        
                        'If glbCompSerial = "S/N - 2395W" Or IsNull(snapEntitle("ED_ANNSICK")) Or snapEntitle("ED_ANNSICK") = 0 Then
                        If (glbCompSerial = "S/N - 2395W") Or (IsNull(snapEntitle("ED_ANNSICK")) Or snapEntitle("ED_ANNSICK") = 0) Or (flgAllowAnnual = True And Not IsNull(snapEntitle("ED_ANNSICK")) And snapEntitle("ED_ANNSICK") <> 0) Then
                        'Do not calculate Annual Sick if it already has value. This is because if Rounded at the time of Sick Ent. Upd
                        'then in this routine it will not do that. Also Ann Sick is calculated at the time of Sick Update.
                            For xRunTimes = 1 To xBalMonths
                                flgAllowAnnual = True
                                
                                blIsLast = False
                                
                                If Not modAnnSick() Then GoTo Final
                                
                                fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                            Next
                        End If
                        
                        snapEntitle.MoveNext
                        
                    Loop 'Wend
                End If
                snapEntitle.Close
            Else
                'Annual Entitlement
                prec = 0
                If Not CR_SnapEntitle("SICK") Then Exit Sub  ' create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount

                        If IsNull(snapEntitle("ED_EFDATES")) = False Then
                            fglbAsOf = snapEntitle("ED_EFDATES")
                        Else
                            GoTo Final
                        End If

                        If Not modAnnSick() Then GoTo Final
 
                        snapEntitle.MoveNext
                    Wend

                End If
                snapEntitle.Close
            End If
            rsEntRules.MoveNext
        Loop 'Until rsEntRules.EOF
        flgAllowAnnual = False
    End If
    
Final:
    Screen.MousePointer = DEFAULT
    If rsEntRules.State <> 0 Then rsEntRules.Close
    Set rsEntRules = Nothing
    Set snapEntitle = Nothing
    
    fbtnCalc = True
    bmk = Data1.Recordset.Bookmark
    Data1.Refresh
    Data1.Recordset.Bookmark = bmk

End Sub

Private Function modAnnSick()
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%
Dim xAsOf
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012
Dim xComments

On Error GoTo modUpdateSelection_Err

modAnnSick = False

'gdbAdoIhr001.BeginTrans

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select
    
    If snapEntitle.State = 0 Then Exit Function
    
    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_ANNSICK")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_ANNSICK")
    End If
    
  
    If IsNull(snapEntitle("ED_PSICK")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PSICK")
    End If
    
    If IsNull(snapEntitle("ED_SICKT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_SICKT")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    'rsJOB.Close    'Ticket #22842 -moved below because of calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    
    If glbLinamar Then dblDHours# = 8
    
    xAsOf = fglbAsOf
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1


    If rsEntRules("VE_EMONTH") > 0 Then
        If dblServiceYears# >= CDbl(rsEntRules("VE_BMONTH")) And dblServiceYears# <= CDbl(rsEntRules("VE_EMONTH")) Then
            intWhereFit& = x%
        End If
    End If

    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    'Ticket #22842 - calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If rsEntRules("VE_TYPE") = "D" Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If rsEntRules("VE_TYPE") = "F" Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close

    dblNewEntitle# = rsEntRules("VE_ENTITLE")
    dblNewMax# = 0
    If rsEntRules("VE_TYPE") = "D" Then           ' Entitlements entered in days
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * xTotEmpHours
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
        Else
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
        End If
        dblEntitleUpd = dblNewEntitle
    End If
    If rsEntRules("VE_TYPE") = "F" Then
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * xTotEmpHours
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
        Else
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
        End If
        dblEntitleUpd = dblNewEntitle
    End If
    If rsEntRules("VE_TYPE") = "H" Then
        If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX")
    End If

    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values

    
    If dblNewMax <> 0 Then          'only do if not zero
        'If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Then 'for town of Ajax or City of Timmins
        If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Or _
            glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2412W" Or glbCompSerial = "S/N - 2399W" Or _
            glbCompSerial = "S/N - 2395W" Then
            
            'for town of Ajax or City of Timmins or St. Leonard's Community Services(Ticket #15071)
            'Ticket #17090 - Township of Wilmot
            'Ticket #17160 - NorWest Community Health Centres
            'Ticket #17111 - West Elgin Community Health Centre
            'Ticket #17315 - The Youth Centre
            
            'The Youth Centre - If they have already earned then subtract that
            If glbCompSerial = "S/N - 2395W" Then
                If (dblEntitleUpd + snapEntitle("ED_SICK") + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                    dblEntitleUpd = dblNewMax - ((dblPrevEntitle# + snapEntitle("ED_SICK")) - dblTKEEntitle#)
                End If
            Else
                If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                    dblEntitleUpd = dblEntitle#
                    
                    'The Youth Centre - If they have already earned then subtract that
                    'If glbCompSerial = "S/N - 2395W" Then
                    '    dblEntitleUpd = dblEntitleUpd - snapEntitle("ED_SICK")
                    'End If
                    
                ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                    dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    
                    'The Youth Centre - If they have already earned then subtract that
                    'If glbCompSerial = "S/N - 2395W" Then
                    '    dblEntitleUpd = dblEntitleUpd - snapEntitle("ED_SICK")
                    'End If
                    
                End If
            End If
        Else
            'Ticket #23878 - KidsLink/Carizon - their Calculated will be Annualized Vacation not using Prev.
            If glbCompSerial = "S/N - 2430W" Then
                If dblEntitleUpd > dblNewMax Then
                    dblEntitleUpd = dblNewMax
                End If
            Else
                If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                    dblEntitleUpd = dblNewMax - dblPrevEntitle#
                End If
            End If
        End If
    End If
    
    If glbCBrant Then
        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
        End If
    End If
    DtTm = Now
    
    xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd

    If dblEntitle# = 0 And glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        snapEntitle("ED_ANNSICK") = IIf(dblEntitleUpd < 0, 0, dblEntitleUpd) '+ IIf(IsNull(snapEntitle("ED_SICK")), 0, snapEntitle("ED_SICK"))
    Else
        If (IsNull(snapEntitle("ED_ANNSICK")) Or snapEntitle("ED_ANNSICK") = 0) Or (UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N") Then
            'Do not calculate Annual Sick if it already has value. This is because if Rounded at the time of Sick Ent. Upd
            'then in this routine it will not do that. Also Ann Sick is calculated at the time of Sick Update.
            snapEntitle("ED_ANNSICK") = dblEntitleUpd
        End If
    End If
    
    snapEntitle.Update
   
lblNextRec:
'gdbAdoIhr001.CommitTrans
    DoEvents



modAnnSick = True


'Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Function CR_SnapEntitle(Optional xType)
Dim SQLQ As String

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

If Not IsMissing(xType) Then
    Call getWSQLQ(xType)
Else
    Call getWSQLQ("")
End If

SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,ED_SICKT,"
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
If Len(rsEntRules("VE_GRPCD")) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & rsEntRules("VE_GRPCD") & "') "
'Ticket #13126 Commented by Frank Jun 5th, 07
'ElseIf glbCompSerial = "S/N - 2376W" Then 'Assembly of First Nations Bryanm 27/Apr/2006 Ticket#10735
'    SQLQ = SQLQ & " AND ED_EMPNBR IN "
'    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'    SQLQ = SQLQ & " WHERE JB_GRPCD <> 'MGT')"
End If

If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub getWSQLQ(Optional xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSECTION
Dim xFromDate
Dim xToDate

fglbESQLQ = glbSeleDeptUn

If Len(rsEntRules("VE_DEPT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & rsEntRules("VE_DEPT") & "' "
If Len(rsEntRules("VE_DIV")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & rsEntRules("VE_DIV") & "' "
If Len(rsEntRules("VE_ORG")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & rsEntRules("VE_ORG") & "' "
If Len(rsEntRules("VE_EMP")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & rsEntRules("VE_EMP") & "' "
If glbLinamar Then
    If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & rsEntRules("VE_SECTION") & "' "
Else
    If Not glbCBrant Then 'added by Bryan 18/Apr/2006 Ticket#10495
        If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & rsEntRules("VE_SECTION") & "' "
    End If
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
    If xType = "VAC" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM1 = '" & rsEntRules("VE_LOC") & "' "
    ElseIf xType = "SICK" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM2 = '" & rsEntRules("VE_LOC") & "' "
    End If
Else
    If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & rsEntRules("VE_LOC") & "' "
End If

If Len(rsEntRules("VE_PT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & rsEntRules("VE_PT") & "' "

If glbCompSerial <> "S/N - 2395W" Then   'Not The Youth Centre - Ticket #17315
    fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
End If

End Sub

Private Function Get_Employee_Last_Accrual_Date(xEmpnbr, xType)
    Dim rsHRAcc As New ADODB.Recordset
    Dim SQLQ As String
    
    'Retrieve Accrual records of the employee based on the xType = VAC or SICK to get the last Accrual update
    'date so that the accrual for the rest of the entitlement period be calculated.
    SQLQ = "SELECT AC_EMPNBR, AC_EDATE FROM HR_ACCRUAL"
    SQLQ = SQLQ & " WHERE AC_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND AC_TYPE = '" & xType & "'"
    SQLQ = SQLQ & " AND AC_COMMENTS NOT LIKE 'Prev.%'"
    SQLQ = SQLQ & " AND AC_ACTION = 'U'"
    SQLQ = SQLQ & " ORDER BY AC_EDATE DESC"
    rsHRAcc.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHRAcc.EOF Then
        rsHRAcc.MoveFirst
        Get_Employee_Last_Accrual_Date = rsHRAcc("AC_EDATE")
    Else
        Get_Employee_Last_Accrual_Date = ""
    End If
    rsHRAcc.Close
    Set rsHRAcc = Nothing

End Function

Private Sub vbxTrueGridVL_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    chkEndDateEdit.Value = False
    
    If fbtnCalc = False And cmdCalc.Caption = "Show Entitlement Dates" Then
        Call ReCalcAnn
        DoEvents
    End If
    Call Display_Value
    fbtnCalc = False

    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
    End If
End Sub

Private Sub vbxTrueGridVLDay_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    chkEndDateEdit.Value = False

    If fbtnCalc = False And cmdCalc.Caption = "Show Entitlement Dates" Then
        Call ReCalcAnn
        DoEvents
    End If
    Call Display_Value
    fbtnCalc = False

    If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
        Call Calc_To_Be_Accrued
    End If
End Sub

Sub cmdOK_Click()
Dim DtTm As Variant, rc As Integer
Dim x, xID
Dim rsHREmp As New ADODB.Recordset
Dim SQLQ As String

DtTm = Now

On Error GoTo Add_Err

If Not chkVac() Then Exit Sub

glbENTScreen = True

'Call UpdUStats(Me) ' update user's stats (who did it and when)

xID = Data1.Recordset("ED_EMPNBR")
If medCalcVac(3).Visible Then 'Hours
    rsDATA!ED_ANNVAC = Val(medCalcVac(3))
ElseIf medCalcVac(2).Visible Then 'Days
    SQLQ = "SELECT ED_EMPNBR, ED_DHRS FROM HREMP WHERE ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        rsDATA!ED_ANNVAC = Val(medCalcVac(2) * IIf(IsNull(rsHREmp("ED_DHRS")), 0, rsHREmp("ED_DHRS")))
    Else
        rsDATA!ED_ANNVAC = 0
    End If
    rsHREmp.Close
End If
'Ticket #14169
If medSickR(3).Visible Then 'Hours
    rsDATA!ED_SICK = Val(medSickR(3))
ElseIf medSickR(2).Visible Then 'Days
    SQLQ = "SELECT ED_EMPNBR, ED_DHRS FROM HREMP WHERE ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        rsDATA!ED_SICK = Val(medSickR(2) * IIf(IsNull(rsHREmp("ED_DHRS")), 0, rsHREmp("ED_DHRS")))
    Else
        rsDATA!ED_SICK = 0
    End If
    rsHREmp.Close
End If
'Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh
Data1.Recordset.Find "ED_EMPNBR=" & xID

'Recalculate for 1 Employee
SQLQ = "ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
Call EntReCalc(SQLQ)
Data1.Refresh
Data1.Recordset.Find "ED_EMPNBR=" & xID

Call SET_UP_MODE

vbxTrueGridVL.EditActive = False
vbxTrueGridVL.MarqueeStyle = 4
vbxTrueGridVL.Enabled = True
vbxTrueGridVL.AllowUpdate = False
vbxTrueGridVLDay.EditActive = False
vbxTrueGridVLDay.MarqueeStyle = 4
vbxTrueGridVLDay.Enabled = True
vbxTrueGridVLDay.AllowUpdate = False


Dim xKey

Call NextForm
Exit Sub

Add_Err:
If Err = 3197 Then Resume Next
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack '28July99 js

Unload Me

End Sub

Private Function chkVac()
Dim dd As Integer

chkVac = False

If medCalcVac(2).Visible And Len(medCalcVac(2)) > 0 Then
    If Not IsNumeric(medCalcVac(2)) Then
        MsgBox "Invalid Annual Vacation Entitlement "
        medCalcVac(2).SetFocus
        Exit Function
    End If
End If
If medCalcVac(3).Visible And Len(medCalcVac(3)) > 0 Then
    If Not IsNumeric(medCalcVac(3)) Then
        MsgBox "Invalid Annual Vacation Entitlement "
        medCalcVac(3).SetFocus
        Exit Function
    End If
End If

chkVac = True

End Function


Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
If rsDATA.State <> 0 Then
    rsDATA.CancelUpdate
End If
Call Display_Value


'Call modSTUPD(True)  ' reset screen's attributes


vbxTrueGridVL.EditActive = False
vbxTrueGridVL.MarqueeStyle = 4
vbxTrueGridVL.Enabled = True
vbxTrueGridVL.AllowUpdate = False
vbxTrueGridVLDay.EditActive = False
vbxTrueGridVLDay.MarqueeStyle = 4
vbxTrueGridVLDay.Enabled = True
vbxTrueGridVLDay.AllowUpdate = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '28July99 js

End Sub

Private Sub Calc_To_Be_Accrued()

    'Vacation - Calculate to be accrued
    'Days
    If medCalcVac(0) <> "" And medCVacDay <> "" Then
        medToBeVac(0) = medCalcVac(0) ' - medCVacDay
        If medToBeVac(0) < 0 Then medToBeVac(0) = 0
    Else
        medToBeVac(0) = ""
    End If
    
    'Hours
    If medCalcVac(1) <> "" And medCVac <> "" Then
        medToBeVac(1) = medCalcVac(1) '- medCVac
        If medToBeVac(1) < 0 Then medToBeVac(1) = 0
    Else
        medToBeVac(1) = ""
    End If
    
    
    'Sick - Calculate to be accrued
    'Days
    If medCalcSick(0) <> "" And medCSickDay <> "" Then
        medToBeSick(0) = medCalcSick(0) '- medCSickDay
        If medToBeSick(0) < 0 Then medToBeSick(0) = 0
    Else
        medToBeSick(0) = ""
    End If
    
    'Hours
    If medCalcSick(1) <> "" And medCSick <> "" Then
        medToBeSick(1) = medCalcSick(1) '- medCSick
        If medToBeSick(1) < 0 Then medToBeSick(1) = 0
    Else
        medToBeSick(1) = ""
    End If

    'Available calculation
    If (medAvailVac(0).Visible = True Or medAvailVac(1).Visible = True) Then
        'If Index = 0 Then
            'Days
            If medVacR(1) = "" Or medToBeVac(0) = "" Then
                medAvailVac(0).Text = ""
            Else
                'medAvailVac(0).Text = (medVacR(1) - medCVacDay) + medCalcVac(0)
                medAvailVac(0).Text = Val(medVacR(1)) + Val(medToBeVac(0))
            End If
            
            If medSickR(1) = "" Or medToBeSick(0) = "" Then
                medAvailSick(0).Text = ""
            Else
                'medAvailSick(0).Text = (medSickR(1) - medCSickDay) + medCalcSick(0)
                medAvailSick(0).Text = Val(medSickR(1)) + Val(medToBeSick(0))
            End If
        'ElseIf Index = 1 Then
            'Hours
            If medVacR(0) = "" Or medToBeVac(1) = "" Then
                medAvailVac(1).Text = ""
            Else
                'medAvailVac(1).Text = (medVacR(0) - medCVac) + medCalcVac(1)
                medAvailVac(1).Text = Val(medVacR(0)) + Val(medToBeVac(1))
            End If
            
            If medSickR(0) = "" Or medToBeSick(1) = "" Then
                medAvailSick(1).Text = ""
            Else
                'medAvailSick(1).Text = (medSickR(0) - medCSick) + medCalcSick(1)
                medAvailSick(1).Text = Val(medSickR(0)) + Val(medToBeSick(1))
            End If
        'End If
    End If


End Sub

Private Sub Hours_Data1_Source()
    Dim SQLQ As String
     SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
     If glbLinamar Then
         SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
         SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
     Else
         If glbOracle Then
             SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
         Else
             SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
         End If
         
     End If
     SQLQ = SQLQ & "ED_EMPNBR,ED_VAC,ED_PVAC,"
     SQLQ = SQLQ & "ED_SICK,ED_PSICK,ED_VACT,ED_SICKT,ED_ANNVAC, ED_ANNSICK, "
     SQLQ = SQLQ & "ED_EFDATE,ED_EFDATES,ED_ETDATE,ED_ETDATES,"
     SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,"
     SQLQ = SQLQ & "ED_PT,ED_ORG,"
     If glbLinamar Then
         SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
         SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
         SQLQ = SQLQ & " ED_VAC/8   AS WK_VACDAY, "
         SQLQ = SQLQ & " ED_PSICK/8 AS WK_PSICKDAY, "
         SQLQ = SQLQ & " ED_SICK/8  AS WK_SICKDAY, "
         SQLQ = SQLQ & " ED_VACT/8  AS WK_VACTDAY, "
         SQLQ = SQLQ & " ED_SICKT/8 AS WK_SICKTDAY, "
         SQLQ = SQLQ & " ED_ANNVAC/8   AS WK_CVDAY, "
         SQLQ = SQLQ & " ED_ANNSICK/8  AS WK_CSDAY, "
         SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
         SQLQ = SQLQ & "(ROUND([ED_VAC]/8,2)+ROUND([ED_PVAC]/8,2)-ROUND([ED_VACT]/8,2)) AS WK_VACODAY, "
         SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
         SQLQ = SQLQ & "(ROUND([ED_PSICK]/8,2)+ROUND([ED_SICK]/8,2)-ROUND([ED_SICKT]/8,2)) AS WK_SICKODAY "
     ElseIf glbOracle Or glbSQL Then
         If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
             
             'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNVAC/ED_DHRS)-(ED_VAC/ED_DHRS) END) AS WK_CVDAY, "
             'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNSICK/ED_DHRS)-(ED_SICK/ED_DHRS) END) AS WK_CSDAY, "
             
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNVAC/ED_DHRS) END) AS WK_CVDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNSICK/ED_DHRS) END) AS WK_CSDAY, "
             
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_VAC,2)+ROUND(ED_PVAC,2))-ROUND(ED_VACT,2)) + ROUND(ED_ANNVAC,2) END) AS WK_AVLVDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_PSICK,2)+ROUND(ED_SICK,2))-ROUND(ED_SICKT,2)) + ROUND(ED_ANNSICK,2) END) AS WK_AVLSDAY, "
             SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
             SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
         Else
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
             SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
             SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
             SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
         End If
     Else
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PVAC]/[ED_DHRS]) AS WK_PVACDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VAC]/[ED_DHRS]) AS WK_VACDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PSICK]/[ED_DHRS]) AS WK_PSICKDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICK]/[ED_DHRS]) AS WK_SICKDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VACT]/[ED_DHRS]) AS WK_VACTDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICKT]/[ED_DHRS]) AS WK_SICKTDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNVAC]/[ED_DHRS]) AS WK_VCDAY, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNSICK]/[ED_DHRS]) AS WK_VSDAY, "
         SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_VAC]+[ED_PVAC]-[ED_VACT])/[ED_DHRS]) AS WK_VACODAY, "
         SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
         SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_PSICK]+[ED_SICK]-[ED_SICKT])/[ED_DHRS]) AS WK_SICKODAY "
     End If
     
     If glbtermopen Then
         SQLQ = SQLQ & ",TERM_SEQ "
         SQLQ = SQLQ & " From Term_HREMP "
     Else
         SQLQ = SQLQ & " From HREMP "
     End If
     SQLQ = SQLQ & "Where " & glbSeleDeptUn
    
     If EESNameSort = True Then
         SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
     Else
         SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
     End If
    
     'SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid1.Columns(ColIndex).DataField & " " & vbxTrueGrid1.Tag
        
    Data1.RecordSource = SQLQ
    Data1.Refresh

End Sub

Private Sub Days_Data1_Source()
    Dim SQLQ As String
    
    SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
    If glbLinamar Then
        SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
        SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
    Else
        If glbOracle Then
            SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
        Else
            SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
        End If
        
    End If
    SQLQ = SQLQ & "ED_EMPNBR,ED_VAC,ED_PVAC,"
    SQLQ = SQLQ & "ED_SICK,ED_PSICK,ED_VACT,ED_SICKT,ED_ANNVAC, ED_ANNSICK, "
    SQLQ = SQLQ & "ED_EFDATE,ED_EFDATES,ED_ETDATE,ED_ETDATES,"
    SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,"
    SQLQ = SQLQ & "ED_PT,ED_ORG,"
    If glbLinamar Then
        SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
        SQLQ = SQLQ & " ED_PVAC/8  AS WK_PVACDAY, "
        SQLQ = SQLQ & " ED_VAC/8   AS WK_VACDAY, "
        SQLQ = SQLQ & " ED_PSICK/8 AS WK_PSICKDAY, "
        SQLQ = SQLQ & " ED_SICK/8  AS WK_SICKDAY, "
        SQLQ = SQLQ & " ED_VACT/8  AS WK_VACTDAY, "
        SQLQ = SQLQ & " ED_SICKT/8 AS WK_SICKTDAY, "
        SQLQ = SQLQ & " ED_ANNVAC/8   AS WK_CVDAY, "
        SQLQ = SQLQ & " ED_ANNSICK/8  AS WK_CSDAY, "
        SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
        SQLQ = SQLQ & "(ROUND([ED_VAC]/8,2)+ROUND([ED_PVAC]/8,2)-ROUND([ED_VACT]/8,2)) AS WK_VACODAY, "
        SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
        SQLQ = SQLQ & "(ROUND([ED_PSICK]/8,2)+ROUND([ED_SICK]/8,2)-ROUND([ED_SICKT]/8,2)) AS WK_SICKODAY "
    ElseIf glbOracle Or glbSQL Then
        If glbCompSerial = "S/N - 2395W" Then   'The Youth Centre - Ticket #17315
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
            
            'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNVAC/ED_DHRS)-(ED_VAC/ED_DHRS) END) AS WK_CVDAY, "
            'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNSICK/ED_DHRS)-(ED_SICK/ED_DHRS) END) AS WK_CSDAY, "
            
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNVAC/ED_DHRS) END) AS WK_CVDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE (ED_ANNSICK/ED_DHRS) END) AS WK_CSDAY, "
            
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2))-ROUND(ED_VACT/ED_DHRS,2)) + ROUND(ED_ANNVAC/ED_DHRS,2) END) AS WK_AVLVDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ((ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2))-ROUND(ED_SICKT/ED_DHRS,2)) + ROUND(ED_ANNSICK/ED_DHRS,2) END) AS WK_AVLSDAY, "
            SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
            SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
        Else
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PVAC/ED_DHRS END) AS WK_PVACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VAC/ED_DHRS END) AS WK_VACDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_PSICK/ED_DHRS END) AS WK_PSICKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICK/ED_DHRS END) AS WK_SICKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_VACT/ED_DHRS END) AS WK_VACTDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_SICKT/ED_DHRS END) AS WK_SICKTDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNVAC/ED_DHRS END) AS WK_CVDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ED_ANNSICK/ED_DHRS END) AS WK_CSDAY, "
            SQLQ = SQLQ & "ED_VAC+ED_PVAC-ED_VACT AS WK_VACO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_VAC/ED_DHRS,2)+ROUND(ED_PVAC/ED_DHRS,2)-ROUND(ED_VACT/ED_DHRS,2) END) AS WK_VACODAY, "
            SQLQ = SQLQ & "ED_PSICK+ED_SICK-ED_SICKT AS WK_SICKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(ED_PSICK/ED_DHRS,2)+ROUND(ED_SICK/ED_DHRS,2)-ROUND(ED_SICKT/ED_DHRS,2) END) AS WK_SICKODAY "
        End If
    Else
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PVAC]/[ED_DHRS]) AS WK_PVACDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VAC]/[ED_DHRS]) AS WK_VACDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_PSICK]/[ED_DHRS]) AS WK_PSICKDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICK]/[ED_DHRS]) AS WK_SICKDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_VACT]/[ED_DHRS]) AS WK_VACTDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_SICKT]/[ED_DHRS]) AS WK_SICKTDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNVAC]/[ED_DHRS]) AS WK_VCDAY, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[ED_ANNSICK]/[ED_DHRS]) AS WK_VSDAY, "
        SQLQ = SQLQ & "[ED_VAC]+[ED_PVAC]-[ED_VACT] AS WK_VACO, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_VAC]+[ED_PVAC]-[ED_VACT])/[ED_DHRS]) AS WK_VACODAY, "
        SQLQ = SQLQ & "[ED_PSICK]+[ED_SICK]-[ED_SICKT] AS WK_SICKO, "
        SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([ED_PSICK]+[ED_SICK]-[ED_SICKT])/[ED_DHRS]) AS WK_SICKODAY "
    End If
    
    If glbtermopen Then
        SQLQ = SQLQ & ",TERM_SEQ "
        SQLQ = SQLQ & " From Term_HREMP "
    Else
        SQLQ = SQLQ & " From HREMP "
    End If
    SQLQ = SQLQ & "Where " & glbSeleDeptUn
   
     If EESNameSort = True Then
         SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
     Else
         SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
     End If
   
    'SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh

End Sub

Private Sub Update_AnnualVac_From_DailyAccrual()
    Dim SQLQ As String
    
    'Get the Annual Vacation for each employee from the Daily Accrual table and update the respective employee's ED_ANNVAC in HREMP.
        
    'Reset Annual Vacation just incase it has changed for this employee
    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0"
    
    'Get Employee's List
    fglbESQLQ = glbSeleDeptUn
    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,ED_SICKT,"
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
    SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
    
    If snapEntitle.State <> 0 Then snapEntitle.Close
    snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If snapEntitle.EOF = False And snapEntitle.BOF = False Then
        'For each employee get Annual Vacation from Daily Accrual file and Update the HREMP
        Do While Not snapEntitle.EOF
            If Not IsNull(snapEntitle("ED_EFDATE")) And Not IsNull(snapEntitle("ED_ETDATE")) Then
                'Get Annual Vacation from the Daily Accrual table and Update the Annual Vacation in HREMP
                snapEntitle("ED_ANNVAC") = Get_AnnualVac_From_DailyAccrual(snapEntitle("ED_EMPNBR"), snapEntitle("ED_ETDATE"))
                snapEntitle("ED_LDATE") = Now
                snapEntitle("ED_LTIME") = Time$
                snapEntitle("ED_LUSER") = glbLEE_ID
                snapEntitle.Update
            Else
                GoTo NextEmp
            End If
NextEmp:
            snapEntitle.MoveNext
        Loop
    End If
    snapEntitle.Close
    Set snapEntitle = Nothing
    
End Sub
