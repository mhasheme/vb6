VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmETERM 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Termination"
   ClientHeight    =   9885
   ClientLeft      =   315
   ClientTop       =   780
   ClientWidth     =   11730
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
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9885
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFrankTest 
      Caption         =   "Frank Test"
      Height          =   375
      Left            =   7800
      TabIndex        =   69
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmWFCBenList 
      Height          =   735
      Left            =   9240
      TabIndex        =   57
      Top             =   8400
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CheckBox chkAllDates 
         Caption         =   "All Date"
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
         Left            =   4560
         TabIndex        =   58
         Top             =   1365
         Width           =   1155
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
         Bindings        =   "feterm.frx":0000
         Height          =   1185
         Left            =   120
         OleObjectBlob   =   "feterm.frx":0014
         TabIndex        =   59
         Top             =   150
         Width           =   10275
      End
      Begin INFOHR_Controls.DateLookup dlpEndDate 
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1800
         TabIndex        =   60
         Tag             =   "41-Effective date of salary change"
         Top             =   1365
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   360
         TabIndex        =   61
         Top             =   1365
         Width           =   885
      End
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   5460
      Left            =   0
      TabIndex        =   28
      Top             =   3120
      Width           =   11475
      _Version        =   65536
      _ExtentX        =   20241
      _ExtentY        =   9631
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
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.Frame fraDiffBenEnd 
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
         Begin Threed.SSCheck chkDiffBenEnd 
            Height          =   225
            Left            =   2100
            TabIndex        =   67
            Top             =   30
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   397
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
         Begin VB.Image imgHelp 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   2500
            Picture         =   "feterm.frx":4C49
            Stretch         =   -1  'True
            Top             =   50
            Width           =   255
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Different Benefit End Dates? "
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
            Index           =   13
            Left            =   0
            TabIndex        =   66
            Tag             =   "41-Benefit End Date"
            Top             =   50
            Width           =   2055
         End
      End
      Begin VB.Frame frmLastDay 
         Height          =   375
         Left            =   3600
         TabIndex        =   62
         Top             =   330
         Visible         =   0   'False
         Width           =   2295
         Begin INFOHR_Controls.DateLookup dlpLastDate 
            Height          =   285
            Left            =   720
            TabIndex        =   63
            Tag             =   "41-Effective date of salary change"
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Day"
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
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   1365
         End
      End
      Begin INFOHR_Controls.DateLookup dlpPosEndDate 
         Height          =   285
         Left            =   3360
         TabIndex        =   13
         Tag             =   "41-Current Position's End Date"
         Top             =   2280
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   330
         Left            =   8040
         TabIndex        =   51
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cboStatFlag3 
         Height          =   315
         Left            =   1920
         TabIndex        =   43
         Top             =   1920
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmdPhoto 
         Appearance      =   0  'Flat
         Caption         =   "&Photo Off"
         Height          =   330
         Left            =   6000
         TabIndex        =   40
         Tag             =   "Print the reports marked with an 'x'"
         Top             =   0
         Width           =   2220
      End
      Begin VB.CommandButton cmdPrintSelected 
         Appearance      =   0  'Flat
         Caption         =   "Print Selected Reports"
         Height          =   330
         Left            =   6000
         TabIndex        =   39
         Tag             =   "Print the reports marked with an 'x'"
         Top             =   450
         Width           =   2220
      End
      Begin VB.ComboBox comDIV 
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
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Tag             =   "41-Division"
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Tag             =   "41-Termination Code "
         Top             =   640
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Tag             =   "41-Date Terminated"
         Top             =   330
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.CommandButton cmdEmail 
         Appearance      =   0  'Flat
         Caption         =   "Send Email"
         Height          =   330
         Left            =   9360
         TabIndex        =   16
         Tag             =   "Print the reports marked with an 'x'"
         Top             =   960
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtComments 
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
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Tag             =   "00-Comments - free form"
         Top             =   3240
         Width           =   8895
      End
      Begin VB.CommandButton cmdTerminate 
         Appearance      =   0  'Flat
         Caption         =   "Terminate the Employee"
         Height          =   330
         Left            =   6000
         TabIndex        =   17
         Tag             =   "Terminate the Employee Selected"
         Top             =   900
         Width           =   2220
      End
      Begin Threed.SSCheck chkRehire 
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Tag             =   "Click to Select Rehire"
         Top             =   2880
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Rehire                                              "
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
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkSum 
         Height          =   225
         Left            =   5760
         TabIndex        =   18
         Tag             =   "Click to Select Summarize Attendance Records    "
         Top             =   2625
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Summarize Attendance Records      "
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
      Begin INFOHR_Controls.DateLookup dlpBenCeaseDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Tag             =   "41-Benefit Cease Date"
         Top             =   1630
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Tag             =   "00-Termination Cause"
         Top             =   0
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TECA"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim2 
         Height          =   285
         Left            =   5690
         TabIndex        =   46
         Top             =   1275
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV2"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   11
         Tag             =   "00-Work Flow Type Code"
         Top             =   4800
         Visible         =   0   'False
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WKFL"
      End
      Begin MSMask.MaskEdBox medAmount 
         Height          =   285
         Left            =   2235
         TabIndex        =   49
         Tag             =   "20- Pensionable Earnings"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "##0.00;(##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.DateLookup dlpDOther2 
         DataSource      =   " "
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Tag             =   "40-Other Date 2"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1045
      End
      Begin Threed.SSCheck chkPosEndDate 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "Click to update Current Positions with End Date"
         Top             =   2280
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Update Current Positions with End Date"
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
         Index           =   3
         Left            =   6720
         TabIndex        =   37
         Tag             =   "00-Enter Union Code"
         Top             =   1680
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin VB.Label lblWFCUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Only complete if the Transfer To Division equals the employee's Division"
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
         Left            =   9120
         TabIndex        =   68
         Top             =   4680
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lblCurDiv 
         AutoSize        =   -1  'True
         Caption         =   "Div"
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
         Left            =   10800
         TabIndex        =   56
         Top             =   2160
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblCurUnion 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   10920
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer to Union"
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
         Left            =   5400
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lbOtherDate2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   53
         Top             =   5160
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Termination"
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
         Height          =   240
         Left            =   6510
         TabIndex        =   52
         Top             =   2325
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgNoSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7710
         Picture         =   "feterm.frx":508B
         Top             =   2325
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deemed PE"
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
         TabIndex        =   50
         Top             =   5160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Flow Type"
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
         TabIndex        =   48
         Top             =   4800
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblVadim2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 2"
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
         Left            =   4040
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Cause"
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
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblTitle 
         Caption         =   "Status Flag 3"
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
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   1950
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit End Date"
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
         TabIndex        =   38
         Tag             =   "41-Benefit End Date"
         Top             =   1635
         Width           =   1815
      End
      Begin VB.Label lblEMPNo 
         AutoSize        =   -1  'True
         Caption         =   "lblEMPNo"
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
         Left            =   2220
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Transfer to Employee#"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Transfer to Facility"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblRehire 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yes"
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
         Left            =   8640
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         TabIndex        =   29
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Tag             =   "41-Date Terminated"
         Top             =   330
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   640
         Width           =   1710
      End
      Begin VB.Image imgSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7680
         Picture         =   "feterm.frx":51D5
         Top             =   2325
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin Threed.SSPanel panTermRpts 
      Height          =   1995
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   5835
      _Version        =   65536
      _ExtentX        =   10292
      _ExtentY        =   3519
      _StockProps     =   15
      Caption         =   "Termination Reports"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Alignment       =   6
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   5
         Tag             =   "Click to select Employee Comments"
         Top             =   1650
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Employee Comments"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Tag             =   "Click to select Compensatory Time"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Compensatory Time "
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   4
         Tag             =   "Click to Select Follow-Ups"
         Top             =   1410
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Follow-Ups"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   3
         Tag             =   "Click to Select Entitlements with Compensatory Time, Hourly Entitlements"
         Top             =   1170
         Width           =   4995
         _Version        =   65536
         _ExtentX        =   8811
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Entitlements with Compensatory Time, Hourly Entitlements"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   2
         Tag             =   "Click to Select Employee Profile"
         Top             =   930
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Employee Profile"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Tag             =   "Click to Select Attendance History - Summarized"
         Top             =   480
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Attendance History - Summarized"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Tag             =   "Click to Select Attendance History - Detail"
         Top             =   720
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Attendance History - Detail"
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
      Begin VB.Label lblRptsPrinted 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Printed for this Employee"
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
         Left            =   3120
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   2340
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11730
      _Version        =   65536
      _ExtentX        =   20690
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7800
         TabIndex        =   44
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   32
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   4440
         TabIndex        =   25
         Top             =   150
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   24
         Top             =   150
         Width           =   720
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   31
      Top             =   9225
      Width           =   11730
      _Version        =   65536
      _ExtentX        =   20690
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6210
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal3 
         Left            =   7560
         Top             =   120
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   10560
      Top             =   8760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Image picPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2625
   End
   Begin VB.Label PicNotF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Photo not Available"
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   6240
      TabIndex        =   41
      Top             =   1440
      Width           =   2115
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmETERM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim fglbEMPNBR As Long
Dim AbortTerm As Boolean
Dim fglbFollowID
Dim fglbNew
Dim glbPicDir, glbPicBMP
Dim locCertNo As String
Dim locWFCPenEligible As Boolean
Dim locWFCPenEarnFlag As Boolean
Dim locSection As String
Dim locUnion As String
Dim locSIN As String
Dim locPayrollID As String
Dim xLocID 'Ticket #23247 Franks 07/22/2013
Dim xLocLastDay 'Ticket #26308 Franks 11/27/2014
Dim xWFCPosChgEmailBody 'Ticket #29343 Franks 10/25/2016
Dim xIsWFCPosChgEmail As Boolean  'Ticket #29343 Franks 10/25/2016

Private Function AUDITTERM()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTACheck As New ADODB.Recordset
Dim xPT As String, xDiv As String, XSNAME As String, xFName As String, xEmpType As String, xDOH As String, xSENDTE As String
Dim SQLQ As String, strFields As String
Dim xAdminBy As String

On Error GoTo AUDIT_ERR

AUDITTERM = False

Dim xBatchID
glbChgTermReason = clpCode(1)
glbChgTermDate = dlpTermDate

If glbTermTran Then 'Ticket #24859 Franks 01/15/2014
    Call AuditFutureDataDele(glbLEE_ID, dlpTermDate.Text)
End If

Call TermPayrollEmp(dlpTermDate, glbLEE_ID, , Termination)
    
rsTB.Open "SELECT ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_DOH,ED_SENDTE,ED_ADMINBY FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If Not IsNull(rsTB("ED_PT")) Then   'Hemu - Gives an error when it's Null and this checking is not done
        xPT = rsTB("ED_PT")
    Else
        xPT = ""
    End If
    
    If Not IsNull(rsTB("ED_DIV")) Then 'George Apr 4,2006
        'xDiv = rsTB("ED_DIV")
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    
    'Ticket #20884 Franks 10/20/2011
    If Not IsNull(rsTB("ED_ADMINBY")) Then
        If IsNull(rsTB("ED_ADMINBY")) Then xAdminBy = "" Else xAdminBy = rsTB("ED_ADMINBY")
    Else
        xAdminBy = ""
    End If
    
    XSNAME = rsTB("ED_SURNAME")
    xFName = rsTB("ED_FNAME")
    If IsNull(rsTB("ED_EMPTYPE")) Then
        xEmpType = ""
    Else
        xEmpType = rsTB("ED_EMPTYPE")
    End If
    xDOH = rsTB("ED_DOH")
    If IsNull(rsTB("ED_SENDTE")) Then
        xSENDTE = ""
    Else
        xSENDTE = rsTB("ED_SENDTE")
    End If
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    xFName = ""
    xEmpType = ""
    xDOH = ""
    xSENDTE = ""
    xAdminBy = ""
End If
rsTB.Close
'Linamar doesn't need Audit records when Transfer Out
'WFC need Audit records when Transfer Out
'Ticket# 7337 For Linamar Interface
'If glbTermTran Or Not glbLinamar Then
    'strFields added by Bryan 02/Dec/05 Ticket#9899
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
    strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_SURNAME, "
    strFields = strFields & "AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID,AU_VADIM2,AU_SIN,AU_SSN,AU_ADMINBY "
    If glbWFC Then 'Ticket #25275 Franks 04/02/2014
        strFields = strFields & ",AU_VSTEP"
    End If
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
    rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
    rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    If glbSamuel Then 'Ticket #20884 Franks 10/20/2011
        rsTA("AU_ADMINBY") = xAdminBy
    End If
    rsTA("AU_EMPTYPE") = xEmpType
    rsTA("AU_SURNAME") = XSNAME
    rsTA("AU_FNAME") = xFName
    rsTA("AU_DOT") = dlpTermDate
    rsTA("AU_TREAS") = clpCode(1)
    'Ticket #16749 dont use Pay Group
    'If glbWFC Then 'Ticket #16616
    '    If Not glbTermTran Then
    '        If Len(clpVadim2.Text) > 0 Then
    '            rsTA("AU_VADIM2") = clpVadim2.Text
    '        End If
    '    End If
    'End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    If glbWFC And glbDivTranInPlant = "Y" Then 'Ticket #25275 Franks 04/02/2014
        rsTA("AU_VSTEP") = "Y"
    End If
    'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_SIN,ED_SSN FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        'Ticket #16749
        If glbWFC Then
            rsTA("AU_SIN") = rsEmp("ED_SIN")
            rsTA("AU_SSN") = rsEmp("ED_SSN")
        End If
    End If
    rsEmp.Close
    'End If
    rsTA.Update
    rsTA.Close
'End If

If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    'Pass Multi Positions to HRAudit for Payweb interface # 7644
    If rsTB.State <> 0 Then rsTB.Close
    rsTB.Open "SELECT JH_EMPNBR,JH_JOB,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    'strFields added by Bryan 02/Dec/05
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
    strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_JOB, AU_ORG, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, "
    strFields = strFields & "AU_LTIME, AU_UPLOAD, AU_TYPE "
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Do While Not rsTB.EOF
        If Not IsNull(rsTB("JH_ORG")) Then
            If Len(rsTB("JH_ORG")) > 0 Then
                '---------------------------
                rsTA.AddNew
                rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
                rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
                rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
                rsTA("AU_EARN_TABL") = "EARN"
                rsTA("AU_NEWEMP") = "N"
                rsTA("AU_JOB") = rsTB("JH_JOB")
                rsTA("AU_ORG") = rsTB("JH_ORG")
                rsTA("AU_DOT") = dlpTermDate
                rsTA("AU_TREAS") = clpCode(1)
                rsTA("AU_COMPNO") = "001"
                rsTA("AU_EMPNBR") = glbLEE_ID
                rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
                rsTA("AU_LUSER") = glbUserID
                rsTA("AU_LTIME") = Time$
                rsTA("AU_UPLOAD") = "N"
                rsTA("AU_TYPE") = "T"
                rsTA.Update
                '---------------------------
            End If
        End If
        rsTB.MoveNext
    Loop
    rsTB.Close
    rsTA.Close
End If

If glbLinamar Or glbWFC Or glbSamuel Then 'For Samuel Ticket #20884 Franks 10/20/2011
    Dim xKey, xCURRENTDIV, xJob
    xKey = "T" & glbTERM_Seq
    rsTB.Open "SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        xJob = rsTB!JH_JOB
    Else
        xJob = ""
    End If
    rsTA.Open "LN_TRALOG", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsTA.AddNew
    
    rsTA!TL_COMPNO = "001"
    rsTA!TL_EMPNBR = glbLEE_ID
    rsTA!TL_SURNAME = XSNAME
    rsTA!TL_FNAME = xFName
    If IsDate(xDOH) Then rsTA!TL_DOH = xDOH
    rsTA!TL_JOB = xJob
    If glbTermTran Then
        rsTA!TL_TYPE = "TERM"
        rsTA!TL_TCOMPLETE = "Y"
        xCURRENTDIV = xDiv
        rsTA!TL_NEWDIV = xCURRENTDIV
        rsTA!TL_NEWEMPNBR = glbLEE_ID
        If Len(xSENDTE) > 0 Then
            rsTA!TL_NEWDIVEDATE = xSENDTE
        End If
        If glbSamuel Then
            rsTA!TL_NEWPLANT = xAdminBy
        End If
    Else
        rsTA!TL_TYPE = "TOUT"
        If glbLinamar Then
            xCURRENTDIV = Left(comDIV, 3)
            rsTA!TL_NEWDIV = xCURRENTDIV
        End If
        If glbWFC Then
            xCURRENTDIV = Trim(Left(comDIV, 4))
            rsTA!TL_NEWDIV = xCURRENTDIV
        End If
        If glbSamuel Then
            xCURRENTDIV = Trim(Left(comDIV, 4))
            rsTA!TL_OLDPLANT = xAdminBy
            rsTA!TL_OLDPLANT_TABL = "EDAB"
            rsTA!TL_NEWPLANT = xCURRENTDIV
            rsTA!TL_NEWPLANT_TABL = "EDAB"
            rsTA!TL_NEWDIV = xDiv
        End If
        
        
        If glbLinamar Then
            'Check if the New Temp. Employee # already exists - if so then change it.
            rsTACheck.Open "SELECT TL_NEWEMPNBR FROM LN_TRALOG WHERE TL_NEWEMPNBR = " & fglbEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsTACheck.EOF Then
                'Look for another #
                Call setNewEmpNbr
                lblEMPNo.Caption = ShowEmpnbr(fglbEMPNBR)
                MsgBox "Please Note: Transfer to Employee # has changed to " & lblEMPNo.Caption
            End If
        End If
        
        'Ticket #21677 Franks 03/14/2012 - union transfer
        If glbWFC Then
            If Len(clpCode(3).Text) > 0 Then
                rsTA!TL_OLD_ORG = lblCurUnion.Caption 'Trim(Mid(lblCurUnion.Caption, InStr(lblCurUnion.Caption, ":") + 2, 4))
                If Len(lblCurUnion.Caption) > 0 Then
                    rsTA!TL_OLD_ORG_DESC = GetTABLDesc("EDOR", lblCurUnion.Caption)
                End If
                rsTA!TL_NEW_ORG = clpCode(3).Text
                rsTA!TL_NEW_ORG_DESC = GetTABLDesc("EDOR", clpCode(3).Text)
            End If
        End If
        
        rsTA!TL_NEWEMPNBR = fglbEMPNBR
        rsTA!TL_NEWDIVEDATE = dlpTermDate
        rsTA!TL_TCOMPLETE = "N"
    End If
    
    rsTA!TL_OLDDIV = xDiv
    rsTA!TL_OLDEMPNBR = glbLEE_ID
    If Len(xSENDTE) > 0 Then
        rsTA!TL_OLDDIVEDATE = xSENDTE
    End If
    rsTA!TL_TOREASON_TABL = "TERM"
    rsTA!TL_TIREASON_TABL = "SDJC"
    rsTA!TL_TOREASON = clpCode(1)
    rsTA!TL_TERM_SEQ = glbTERM_Seq
    
    rsTA!TL_KEY = xKey
    rsTA!TL_CURRENTDIV = xCURRENTDIV
    
    rsTA("TL_LDATE") = Format(Now, "SHORT DATE")
    rsTA("TL_LUSER") = glbUserID
    rsTA("TL_LTIME") = Time$
    rsTA.Update
    rsTA.Close
    rsTA.Open "SELECT TL_KEY,TL_CURRENTDIV FROM LN_TRALOG WHERE TL_KEY='E" & glbLEE_ID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do Until rsTA.EOF
        rsTA!TL_KEY = xKey
        rsTA!TL_CURRENTDIV = xCURRENTDIV
        rsTA.Update
        rsTA.MoveNext
    Loop
    rsTA.Close
'    gdbAdoIhr001.Execute "UPDATE LN_TRALOG SET TL_KEY='" & xKEY & "' WHERE TL_KEY='E" & glbLEE_ID & "'"

End If

AUDITTERM = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js
Resume Next
End Function

Private Function AUDITCOUNSEL()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsCounsel As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITCOUNSEL = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
rsTB.Close
Set rsTB = Nothing

rsCounsel.Open "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsCounsel.EOF Then
    'No Counseling data
    AUDITCOUNSEL = True
    
    rsCounsel.Close
    Set rsCounsel = Nothing
    Exit Function
End If

strFields = "AU_TYPE_TABL, AU_REASON_TABL, AU_ATTREASON_TABL, "
strFields = strFields & "AU_PTUPL, AU_DIVUPL, AU_COMMENTS, AU_COUBY, AU_COMPLETED, AU_EMP_RESPONSE, "
strFields = strFields & "AU_TYPE, AU_REASON, AU_ATTREASON, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TRANS_TYPE, "
strFields = strFields & "AU_COUDATE, AU_INCDATE, AU_FOLLOWUPD1, AU_FOLLOWUPD2, AU_FOLLOWUPD3, AU_ATTDATE, AU_DATE1, AU_EMP_AGREED, AU_EMP_DECLINED "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT_COUNSEL WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

Do While Not rsCounsel.EOF
    rsTA.AddNew
    rsTA("AU_TYPE_TABL") = "CETY": rsTA("AU_REASON_TABL") = "CERE": rsTA("AU_ATTREASON_TABL") = "ADRE"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    
    rsTA("AU_TYPE") = rsCounsel("CL_TYPE")
    rsTA("AU_COUDATE") = rsCounsel("CL_COUDATE")
    If Not IsNull(rsCounsel("CL_REASON")) Then rsTA("AU_REASON") = rsCounsel("CL_REASON")
    If Not IsNull(rsCounsel("CL_COUBY")) Then rsTA("AU_COUBY") = rsCounsel("CL_COUBY")
    If Not IsNull(rsCounsel("CL_INCDATE")) Then rsTA("AU_INCDATE") = rsCounsel("CL_INCDATE")
    If Not IsNull(rsCounsel("CL_EMP_AGREED")) Then rsTA("AU_EMP_AGREED") = rsCounsel("CL_EMP_AGREED")
    If Not IsNull(rsCounsel("CL_EMP_DECLINED")) Then rsTA("AU_EMP_DECLINED") = rsCounsel("CL_EMP_DECLINED")
    If Not IsNull(rsCounsel("CL_COMMENTS")) Then rsTA("AU_COMMENTS") = rsCounsel("CL_COMMENTS")
    If Not IsNull(rsCounsel("CL_EMP_RESPONSE")) Then rsTA("AU_EMP_RESPONSE") = rsCounsel("CL_EMP_RESPONSE")
    
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TRANS_TYPE") = "T"
    rsTA.Update
    
    rsCounsel.MoveNext
Loop
rsTA.Close
Set rsTA = Nothing

rsCounsel.Close
Set rsCounsel = Nothing

AUDITCOUNSEL = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js

End Function

Private Sub chkDiffBenEnd_Click(Value As Integer) 'Ticket #27245 Franks 09/02/2015 non WFC Diff Benefit End Date
    If chkDiffBenEnd.Value Then
        lblTitle(5).Enabled = False
        dlpBenCeaseDate.Enabled = False
        dlpBenCeaseDate.Text = ""
        Call WFCBenListScreen(glbLEE_ID)
    Else
        lblTitle(5).Enabled = True
        dlpBenCeaseDate.Enabled = True
        frmWFCBenList.Visible = Value
    End If
    
End Sub

Private Sub chkPosEndDate_Click(Value As Integer)
    '8.0 - Ticket #22682 - Update Current Positions with End Date
    If chkPosEndDate Then
        If IsDate(dlpTermDate.Text) And Not IsDate(dlpPosEndDate.Text) Then
            dlpPosEndDate.Text = dlpTermDate.Text
            dlpPosEndDate.Enabled = True
        End If
    Else
        dlpPosEndDate.Text = ""
        dlpPosEndDate.Enabled = False
    End If
End Sub

Private Sub chkPosEndDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkRehire_Click(Value As Integer)

If chkRehire.Value = True Then
    lblRehire.Caption = "Yes"
Else
    lblRehire.Caption = "No"
End If

End Sub

Private Sub chkRehire_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub chkSum_Click(Value As Integer)

If chkSum.Value = False Then
    glbchkSum = False
Else
    glbchkSum = True
End If

End Sub

Private Function chkTerms()
Dim dd As Integer
Dim xFromCountry As String
Dim xToCountry As String
Dim locDiv As String
Dim locEmpStatus As String
Dim Msg$, DgDef As Variant, Response%, Title$ 'Ticket #23611 Franks 05/14/2013
Dim SQLQ As String 'Ticket #23247 Franks 07/23/2013
Dim rsTemp As New ADODB.Recordset
Dim locBenGroup As String 'Ticket #24176 Franks 08/07/2013

chkTerms = False

If Len(dlpTermDate.Text) < 1 Then
    MsgBox TranStr("Termination Date is a required field")
    dlpTermDate.SetFocus
    Exit Function
End If

If Not IsDate(dlpTermDate.Text) Then
    MsgBox TranStr("Termination Date is not a valid date.")
    dlpTermDate.SetFocus
    Exit Function
End If

locCertNo = ""
locDiv = ""
locEmpStatus = ""
locSection = ""
locUnion = ""
locSIN = ""
locPayrollID = ""
locBenGroup = "" 'Ticket #24176 Franks 08/07/2013
If IsDate(dlpTermDate.Text) Then
    Dim rsEM As New ADODB.Recordset
    
    rsEM.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
        If Not IsNull(rsEM("ED_DOH")) Then 'ticket# 8617
            If DaysBetween(rsEM("ED_DOH"), dlpTermDate) < 0 Then
                MsgBox "Termination Date must be greater than Original Hire Date"
                rsEM.Close
                dlpTermDate.SetFocus
                Exit Function
            End If
        Else
            MsgBox "Original Hire Date cannot be blank"
            rsEM.Close
            Exit Function
        End If
        If Not IsNull(rsEM("ED_USER_TEXT1")) Then
            locCertNo = rsEM("ED_USER_TEXT1")
        End If
        If Not IsNull(rsEM("ED_DIV")) Then
            locDiv = rsEM("ED_DIV")
        End If
        If Not IsNull(rsEM("ED_EMP")) Then
            locEmpStatus = rsEM("ED_EMP")
        End If
        If Not IsNull(rsEM("ED_SECTION")) Then
            locSection = rsEM("ED_SECTION")
        End If
        If Not IsNull(rsEM("ED_ORG")) Then
            locUnion = rsEM("ED_ORG")
        End If
        If Not IsNull(rsEM("ED_SIN")) Then
            locSIN = rsEM("ED_SIN")
        End If
        If Not IsNull(rsEM("ED_PAYROLL_ID")) Then
            locPayrollID = rsEM("ED_PAYROLL_ID")
        End If
        If Not IsNull(rsEM("ED_BENEFIT_GROUP")) Then locBenGroup = rsEM("ED_BENEFIT_GROUP") 'Ticket #24176 Franks 08/07/2013
    rsEM.Close
End If

If Len(clpCode(1).Text) < 1 Or clpCode(1).Caption = "Unassigned" Then
    MsgBox TranStr("Termination Reason is a required field")
    clpCode(1).SetFocus
    Exit Function
End If

If glbTermTran Then
    If glbLinamar Or glbWFC Then
        If clpCode(1) = "TOUT" Then
            MsgBox "The reason code TOUT is not allowed. Please use Transfer Out function to transfer the employee"
            clpCode(1).SetFocus
            Exit Function
        End If
        If clpCode(1) = "TRA" Then
            MsgBox "The reason code TRA is not allowed. "
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
Else
    'Ticket #29660 - No Transfer Out for CONP employees
    If glbWFC Then
        If locEmpStatus = "CONP" Then
            Msg$ = "Contractor (CONP) cannot be Transferred Out."
            MsgBox Msg$, vbOKOnly, "info:HR - Contractual Employees"
            Exit Function
        End If
    End If
    
    If comDIV.ListIndex = -1 Then
        MsgBox lStr("Invalid Division")
        comDIV.SetFocus
        Exit Function
    End If
    'Ticket #16749  dont use Pay Group
    'If glbWFC Then 'Ticket #16748
    '    If glbEmpCountry = "U.S.A." Then
    '        If Len(clpVadim2.Text) = 0 Then
    '            MsgBox lStr("Vadim Field 2") & " is a required field"
    '            clpVadim2.SetFocus
    '            Exit Function
    '        End If
    '    End If
    'End If
    
    'Ticket #18654
    'On Transfer Out, the Benefit End Date is mandatory if the To Division's Division Master
    'Country is not equal to "CANADA" when the From Division's country is "CANADA"
    If glbWFC Then
        'Ticket #27827 Franks 12/14/2015 - begin
        If Not glbDivTranInPlant = "Y" Then
            If Len(clpCode(3).Text) = 0 Then 'not Union Transfer
            '"   If the Transfer To Division equals the employee's plant code, display a message saying "Do not use this transfer program if the employee is only transferring to another Division within the same plant. Use the "Transfer Division within Plant" function.".
                If Len(Left(comDIV.Text, 4)) > 0 Then
                    If getSectionByDiv(Left(comDIV.Text, 4)) = getSectionByDiv(lblCurDiv.Caption) Then
                        Msg$ = "Do not use this transfer program if the employee is only transferring to another Division within the same plant. Use the 'Transfer Division within Plant' function. "
                        MsgBox Msg$
                        clpCode(3).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
        'Ticket #27827 Franks 12/14/2015 - end
        
        'Ticket #21677 Franks 03/14/2012 - begin
        'Union Transfer
        'If union enter then the division must be the current. the union transfer only in the same Div
        'Union can't be the same
        If Len(clpCode(3).Text) > 0 Then
            If clpCode(3).Caption = "Unassigned" Then
                MsgBox lStr("Union") & " code must be valid."
                clpCode(3).SetFocus
                Exit Function
            End If
            If Not Left(comDIV.Text, 4) = lblCurDiv.Caption Then
                MsgBox lStr("Division") & " must be " & lblCurDiv.Caption & " for " & lStr("Union") & " Transfer."
                clpCode(3).SetFocus
                'Ticket #27827 Franks 12/14/2015 - If the Union is entered, and the Transfer To Division does not equal the employee's Division, clear the Transfer to Union data and display the message above
                clpCode(3).Text = ""
                Exit Function
            End If
            If clpCode(3).Text = lblCurUnion.Caption Then
                MsgBox "Invalid " & lStr("Union") & " code. The current code is " & clpCode(3).Text
                clpCode(3).SetFocus
                Exit Function
            End If
        End If
        'Ticket #21677 Franks 03/14/2012 - end
        If Len(locDiv) > 0 Then
            If Len(dlpBenCeaseDate.Text) = 0 Then
                xFromCountry = GetCountryFromDiv(locDiv)
                If xFromCountry = "CANADA" And dlpBenCeaseDate.Enabled Then 'Ticket #19955 Franks 03/07/2011
                    xToCountry = GetCountryFromDiv(Left(comDIV.Text, 4))
                    If Not xToCountry = "CANADA" Then
                        MsgBox "Benefit End Date is required! " & Chr(10) & "This employee is going to be transferred from " & xFromCountry & " to " & xToCountry
                        dlpBenCeaseDate.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
End If

'Grey out Work Flow Master and Detail. Remove the Work Flow code on Termination.
'December 4, 2009 on WFC Pension Outstanding Tasks By Dec1009.doc
'If glbWFC Then
'    If locWFCPenEligible Then
'        If Len(clpCode(0).Text) = 0 Then
'            MsgBox "Work Flow Type is reqired because " & lStr("Employment Type") & " is 'Yes'"'
'            clpCode(0).SetFocus
'            Exit Function
'        End If
'    End If
'End If

If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21518 Franks 05/03/2012
    If clpCode(1).Text = "RETI" Then
        If Len(dlpBenCeaseDate.Text) = 0 Then
            MsgBox "Benefit End Date is required if Termination Reason is 'RETI'. " & Chr(10) & "Default it to Termination Date"
            dlpBenCeaseDate.Text = dlpTermDate.Text
            dlpBenCeaseDate.SetFocus
            Exit Function
        End If
    End If
End If

If glbWFC Then
    '07/12/2010 turn on this function again - "Pension Changes - June 1-2010.docx"
    '02/25/10 from Jerry & Margaret email
    'If Country of Employment = "CANADA", do not allow "RET" to be used as a termination reason.
    If locWFCPenEligible Then
        'Ticket #18331 turn off this for now, this will turn on wheneve RET transaction function is done
        If clpCode(1).Text = "RET" Then
            'MsgBox "Cannot enter 'RET' as Termination Reason for Canadian employees."
            MsgBox "Transaction cannot be processed." & Chr(10) & " Please use the Retirement function."
            'clpCode(1).SetFocus
            Exit Function
        End If
        
        'Ticket #21021 Franks 02/29/2012
        '6.  On termination - if the reason is ERIN the cause must be entered, and cause description needs
        'to print in the pension e-mail but not the termination email.
        'Also, print the Home Telephone Number in the Pension email.
        If glbTermTran Then  'for Termination
            'If clpCode(1).Text = "ERIN" Then 'Ticket #27820 Franks 11/25/2015 - this field is required for all employees
            If clpCode(2).Visible Then
                If Len(clpCode(2).Text) = 0 Then
                    'MsgBox "Termination Cause is required if the Termination Reason is ERIN."
                    MsgBox "Termination Cause is required " 'if the Termination Reason is ERIN."
                    clpCode(2).SetFocus
                    Exit Function
                End If
            End If
            'End If
        End If
    End If
    If locWFCPenEarnFlag Then 'Ticket #18265
        If Len(medAmount.Text) = 0 Then
            MsgBox "Deemed PE must be entered since Union is 'NONE' and" & Chr(10) & "No 'EN01' Payroll Transaction record is found for current year."
            Exit Function
        Else
            If Not IsNumeric(medAmount.Text) Then
                MsgBox "Invalid Deemed PE"
                Exit Function
            End If
        End If
        'If Employment Status is a LOA Status, display a message "This employee is in
        'a non Active Employment Status. The Deemed PE may need be manually calculated instead of using the Pensionable Earnings identified by payroll."
        'see Pension Tests - April0810.docx
        If isEmpLOA(glbLEE_ID) Then
            'MsgBox "A non Active Employment Status. The Deemed PE may need be manually calculated" & Chr(10) & "instead of using the Pensionable Earnings identified by payroll."
            MsgBox "This employee has a leave of absence status. The Deemed PE on the PA Master" & Chr(10) & "may need to be confirmed. Some LOA statuses are pensionable."
        End If
        If locEmpStatus = "RET" And clpCode(1).Text = "DECD" Then
                MsgBox "Cannot use Reason Code DECD when the Employment Status equals RET." & Chr(10) & " Please use the Death of a Retiree function."
                Exit Function
        End If
    End If

    'Ticket #19266 Franks 12/02/2010
    If dlpDOther2.Visible Then
        If Not glbTermTran Then 'transfer out
            'Ticket #24936 Franks 02/05/2014
            'If it is a division transfer, the NGS End Date is not required. The NGS End Date is only required if it's a Union transfer
            If Len(clpCode(3).Text) > 0 Then
                If Not IsDate(dlpDOther2.Text) Then
                    MsgBox lStr("Other Date 2") & " is required field"
                    dlpDOther2.SetFocus
                    Exit Function
                End If
            End If
        Else
            If Not IsDate(dlpDOther2.Text) Then
                MsgBox lStr("Other Date 2") & " is required field"
                dlpDOther2.SetFocus
                Exit Function
            End If
        End If
        'Ticket #23611 Franks 05/14/2013
        If Not locBenGroup = "13" Then   'Ticket #24176 Franks 08/07/2013
            'If Not CVDate(dlpTermDate.Text) = CVDate(dlpDOther2.Text) Then
            If Not (dlpTermDate.Text = dlpDOther2.Text) Then
                Msg$ = "Termination Date and NGS End Date do not match." & Chr(10) & "Is this okay? "
                Title$ = TranStr("Terminate Employee")
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDNO Then    ' Evaluate response
                    dlpDOther2.Text = dlpTermDate.Text
                    dlpDOther2.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If glbWFC And frmWFCBenList.Visible Then 'Ticket #23247 Franks 07/23/2013
        'check if there are benefits
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "' AND NOT (BM_ENDDATE IS NULL) "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            'For Company % Only
            SQLQ = "SELECT * FROM HRBENGRPLIST "
            SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "' AND NOT (BM_ENDDATE IS NULL) "
            SQLQ = SQLQ & "AND BM_PCC = 1 " 'NEW - Company % only
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsTemp.EOF Then 'no End Date enter, then pop up this message
                Msg$ = "Do the Company Paid benefits have the same End Date as the NGS End Date?"
                Title$ = "Employee Termination"
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                
                If Response% = IDNO Then    ' Evaluate response
                    Exit Function 'If No, do not terminate the employee. They will manually enter the end dates on the benefits.
                Else 'Yes
                    'If Yes, automatically enter the End Date for all company paid benefits
                    Call UptData2fromDOT
                    Exit Function
                End If
            End If
        End If
        If rsTemp.State <> 0 Then rsTemp.Close
    End If
End If

If Len(dlpBenCeaseDate.Text) > 0 Then
    If Not IsDate(dlpBenCeaseDate.Text) Then
        MsgBox "Benefit End Date is not a valid date."
        dlpBenCeaseDate.SetFocus
        Exit Function
    End If
End If

'8.0 - Ticket #22682 - Update Current Positions with End Date.
If glbTermTran Then
    If chkPosEndDate Then
        If Len(Trim(dlpPosEndDate.Text)) = 0 Then
            MsgBox "Current Position End Date is required"
            dlpPosEndDate.SetFocus
            Exit Function
        ElseIf Not IsDate(dlpPosEndDate.Text) Then
            MsgBox "Current Position End Date is invalid."
            dlpPosEndDate.SetFocus
            Exit Function
        End If
    End If
End If

chkTerms = True

End Function

Function EERetrieve()
    Call cll_EEFind(Me)
End Function

Private Sub cll_EEFind(frmName As Form)

    frmName.Enabled = True
    frmName.lblEENum = ShowEmpnbr(glbLEE_ID)
    frmName.lblEEID = glbLEE_ID
    frmName.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    If glbTermTran Then
        frmName.Caption = "Termination - " & Left$(lblEEName, 100)
    Else
        If glbDivTranInPlant = "Y" Then
            frmName.Caption = "Transfer Division within Plant - " & Left$(lblEEName, 100)
        Else
            frmName.Caption = "Transfer Out - " & Left$(lblEEName, 100)
        End If
    End If
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
    
    If glbWFC Then
        locWFCPenEligible = False
        locWFCPenEarnFlag = False
        Call setPenEarningsBox(False) 'reset to invisible
        If glbTermTran Then
            locWFCPenEligible = WFCPensionEligible(glbLEE_ID)
            Call WFCDeemedPEsetup(Date)
            
            'Ticket #19266 Franks 12/02/10
            'On Termination screen, add "Other Date 2" to be completed by the user. Not optional
            Call WFCOther2Screen(glbLEE_ID)
        Else
            'Ticket #19955 Franks 03/07/2011
            'Transfer Out screen:"   If County of Employment is not CANADA, grey out "Benefit End Date".
            lblTitle(5).Enabled = False
            dlpBenCeaseDate.Text = ""
            dlpBenCeaseDate.Enabled = False
            
            If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/17/2014
                'Div transfer in same plant: no affect to NGS, Benefit and MLF
            Else
                'Ticket #24767 Franks 12/11/2013
                Call WFCOther2Screen(glbLEE_ID)
            End If
            
        End If
    End If
    
    'Ticket #16616
    'Get Pay Group (Vadim 2)
    'Ticket #16749 dont use Pay Group
    'If glbWFC Then
    '    If Not glbTermTran Then
    '        clpVadim2.Text = "" ' GetEmpData(glbLEE_ID, "ED_VADIM2")
    '    End If
    'End If

End Sub

Private Sub chkSum_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkTermRpts_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    If glbWFC Then
        If Index = 2 Then
            clpCode(2).TransDiv = GetTransDivTReason(clpCode(1).Text)
        End If
    End If
End Sub

Public Sub cmdEmailWFCPension()
    Dim rsPen As New ADODB.Recordset
    Dim rsLocEmp As New ADODB.Recordset
    Dim MailBody As String
    Dim SecCode As String, SecDesc As String
    Dim PenType As String
    Dim UnionCode As String
    Dim SalHrl As String
    Dim xEmpNo As Double
    Dim SQLQ As String
    Dim xStr As String
    Dim xTmpVal As Double
    Dim xCredSer As Double
    Dim xContSer As Double
    Dim DBEarns, DBCR, DBCS, DBCalDB, DBCashout
    Dim DBEarnsHly, DBCRHly, DBCSHly, DBCalDBHly, DBCashoutHly
    Dim DCEarns, DCER, DCEE, DCCashout
    Dim xSalFlag As Boolean
    Dim xHlyFlag As Boolean
    Dim xDBList As String
    Dim xTermReason As String
    Dim xtmpStr As String
    Dim xAlias As String 'Ticket #27983 Franks 02/09/2016
    
    On Error GoTo ErrorHandler
    
    If glbDivTranInPlant = "Y" Then 'Ticket #25307 Franks 04/08/2014 - don't send email during transfer within plants
        Exit Sub
    End If
    
    xDBList = "AND (LEFT(PE_PENSIONTYPE,2) = 'DB' OR LEFT(PE_PENSIONTYPE,3) = 'IDL' OR LEFT(PE_PENSIONTYPE,3) = 'UPG' OR LEFT(PE_PENSIONTYPE,3) = 'PRE' OR PE_PENSIONTYPE = 'DBSUP' OR PE_PENSIONTYPE = 'MON' ) "

    'Exit Sub
    Load frmSendEmail
    If glbWFC And clpCode(1) = "TOUT" Then   'Ticket #23173 Franks 01/28/2013 - for transfer out
        frmSendEmail.txtSubject.Text = "info:HR Transfer Notice - " & lblEEName.Caption
    Else
        'frmSendEmail.txtSubject.Text = "info:HR Pension Termination Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Termination Notice - " & lblEEName.Caption
    End If
    'Ticket #27983 Franks 02/09/2016 - begin
    xAlias = GetEmpData(lblEENum.Caption, "ED_ALIAS")
    If Len(xAlias) > 0 Then
        frmSendEmail.txtSubject.Text = frmSendEmail.txtSubject.Text & "(" & xAlias & ")"
    End If
    
    If glbWFC And clpCode(1) = "TOUT" Then   'Ticket #23173 Franks 01/28/2013 - for transfer out
        MailBody = "The employee below has been transferred out." & vbCrLf & vbCrLf
        MailBody = MailBody & "Transfer Out Information:" & vbCrLf 'Ticket #20139
    Else
        MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
        MailBody = MailBody & "Termination Information:" & vbCrLf 'Ticket #20139
    End If
    'MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
    'MailBody = MailBody & "Termination Information:" & vbCrLf 'Ticket #20139
    'Ticket #27983 Franks 02/09/2016 - end
    MailBody = MailBody & Space(4) & "Employee #: " & lblEENum.Caption & vbCrLf
    If Len(xAlias) > 0 Then
        MailBody = MailBody & Space(4) & "Name: " & lblEEName.Caption & "(" & xAlias & ")" & vbCrLf
    Else
        MailBody = MailBody & Space(4) & "Name: " & lblEEName.Caption & vbCrLf
    End If
    frmSendEmail.txtTo.Text = "pension@woodbridgegroup.com"
    frmSendEmail.txtCC.Text = GetCurUserEmail 'Ticket #19852 Franks 02/14/2011
    
    xEmpNo = lblEENum.Caption
    SecCode = GetEmpData(xEmpNo, "ED_SECTION")
    UnionCode = GetEmpData(xEmpNo, "ED_ORG")
    SecDesc = GetTABLDesc("EDSE", SecCode)
    SalHrl = GetSalHourly(SecCode, UnionCode)
    PenType = GetPensionType(SecCode, UnionCode)
    xTermReason = clpCode(1).Caption
    
    MailBody = MailBody & Space(4) & lStr("Section: ") & "" & SecDesc & vbCrLf
    MailBody = MailBody & Space(4) & "Salaried/Hourly: " & SalHrl & vbCrLf

    MailBody = MailBody & Space(4) & "Termination Date: " & dlpTermDate.Text & vbCrLf
    MailBody = MailBody & Space(4) & "Reason: " & xTermReason & vbCrLf
    If Len(clpCode(2).Text) > 0 Then
        MailBody = MailBody & Space(4) & "Cause: " & clpCode(2).Caption & vbCrLf
    End If
    MailBody = MailBody & vbCrLf
    
    'Ticket #20139 Franks 04/11/2011 - begin
    MailBody = MailBody & "Personal Information:" & vbCrLf 'Ticket #20139
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    SQLQ = "SELECT ED_EMPNBR, ED_DOB, ED_DOH, ED_MSTAT,ED_ADDR1, ED_ADDR2,ED_CITY,ED_PROV,ED_PCODE,ED_COUNTRY,ED_PHONE FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        If Not IsNull(rsLocEmp("ED_DOB")) Then
            MailBody = MailBody & Space(4) & "Date of Birth: " & rsLocEmp("ED_DOB") & vbCrLf
        End If
        If Not IsNull(rsLocEmp("ED_DOH")) Then
            MailBody = MailBody & Space(4) & "Date of Hire: " & rsLocEmp("ED_DOH") & vbCrLf
        End If
        If Not IsNull(rsLocEmp("ED_MSTAT")) Then
            xtmpStr = getMSDesc(rsLocEmp("ED_MSTAT"))
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Marital Status: " & xtmpStr & vbCrLf
            End If
        End If
        If Not IsNull(rsLocEmp("ED_ADDR1")) Then
            xtmpStr = rsLocEmp("ED_ADDR1")
            If Not IsNull(rsLocEmp("ED_ADDR2")) Then
                If Len(rsLocEmp("ED_ADDR2")) > 0 Then
                    xtmpStr = xtmpStr & " " & rsLocEmp("ED_ADDR2")
                End If
            End If
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Address: " & xtmpStr & vbCrLf
            End If
        End If
        If Not IsNull(rsLocEmp("ED_CITY")) Then
            xtmpStr = rsLocEmp("ED_CITY")
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "City: " & xtmpStr & vbCrLf
            End If
        End If
        If Not IsNull(rsLocEmp("ED_PROV")) Then
            xtmpStr = rsLocEmp("ED_PROV")
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Province: " & xtmpStr & vbCrLf
            End If
        End If
        If Not IsNull(rsLocEmp("ED_PCODE")) Then
            xtmpStr = rsLocEmp("ED_PCODE")
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Postal Code: " & xtmpStr & vbCrLf
            End If
        End If
        If Not IsNull(rsLocEmp("ED_COUNTRY")) Then
            xtmpStr = rsLocEmp("ED_COUNTRY")
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Country: " & xtmpStr & vbCrLf
            End If
        End If
        'Ticket #21021 Franks 02/29/2012 - print the Home Telephone Number in the Pension email
        If Not IsNull(rsLocEmp("ED_PHONE")) Then
            xtmpStr = rsLocEmp("ED_PHONE")
            If Len(xtmpStr) > 0 Then
            MailBody = MailBody & Space(4) & "Telephone: " & xtmpStr & vbCrLf
            End If
        End If
        
    End If
    rsLocEmp.Close
    MailBody = MailBody & vbCrLf
    'Ticket #20139 Franks 04/11/2011 - end
    
    'DB Pensions - Salaried
    xSalFlag = False
    DBEarns = 0: DBCR = 0: DBCS = 0: DBCalDB = 0: DBCashout = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & glbSIN & "' "
    SQLQ = SQLQ & xDBList
    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Salaried' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        xSalFlag = True
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DBEarns = DBEarns + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_CREDITED_SERV")) Then
            DBCR = DBCR + rsPen("PE_CREDITED_SERV")
        End If
        If Not IsNull(rsPen("PE_CONT_SERV")) Then
            DBCS = DBCS + rsPen("PE_CONT_SERV")
        End If
        If Not IsNull(rsPen("PE_ANNDEFERRED")) Then
            DBCalDB = DBCalDB + rsPen("PE_ANNDEFERRED")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DBCashout = DBCashout + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    'If DBEarns > 0 Then
    If xSalFlag Then
        xStr = "Salaried DB Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Earned Pension: " & "$" & DBEarns
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Credited Service: " & Format((DBCR / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cont. Service: " & Format((DBCS / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Calculated Pension: " & "$" & DBCalDB
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DBCashout
        MailBody = MailBody & xStr & vbCrLf & vbCrLf
    End If
    
    'DB Pensions - Hourly
    xHlyFlag = False
    DBEarnsHly = 0: DBCRHly = 0: DBCSHly = 0: DBCalDBHly = 0: DBCashoutHly = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & glbSIN & "' "
    SQLQ = SQLQ & xDBList
    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        xHlyFlag = True
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DBEarnsHly = DBEarnsHly + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_CREDITED_SERV")) Then
            DBCRHly = DBCRHly + rsPen("PE_CREDITED_SERV")
        End If
        If Not IsNull(rsPen("PE_CONT_SERV")) Then
            DBCSHly = DBCSHly + rsPen("PE_CONT_SERV")
        End If
        If Not IsNull(rsPen("PE_ANNDEFERRED")) Then
            DBCalDBHly = DBCalDBHly + rsPen("PE_ANNDEFERRED")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DBCashoutHly = DBCashoutHly + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    'If DBEarnsHly > 0 Then
    If xHlyFlag Then
        xStr = "Hourly DB Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Earned Pension: " & "$" & DBEarnsHly
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Credited Service: " & Format((DBCRHly / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cont. Service: " & Format((DBCSHly / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Calculated Pension: " & "$" & DBCalDBHly
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DBCashoutHly
        MailBody = MailBody & xStr & vbCrLf & vbCrLf
    End If
    
    'DC Pension
    DCEarns = 0: DCER = 0: DCEE = 0: DCCashout = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_EMPNBR = " & xEmpNo & " "
    If Len(SecCode) > 0 Then
        SQLQ = SQLQ & "AND PE_SECTION = '" & SecCode & "' "
    End If
    SQLQ = SQLQ & "AND PE_PENSIONTYPE = 'DC' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DCER = DCER + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_MEM_DOLLAR")) Then
            DCEE = DCEE + rsPen("PE_MEM_DOLLAR")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DCCashout = DCCashout + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    If DCER > 0 Then
        xStr = "DC Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Employer Portion: " & "$" & DCER
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Employee Portion: " & "$" & DCEE
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DCCashout
        MailBody = MailBody & xStr & vbCrLf
    End If
    
    'Ticket #20139 Franks 04/11/2011 - begin
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    SQLQ = "SELECT * FROM HRBENS "
    SQLQ = SQLQ & " WHERE BD_EMPNBR = " & xEmpNo & " AND BD_BCODE = 'DB'"
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        MailBody = MailBody & "Beneficiary Information:" & vbCrLf 'Ticket #20139
        If Not IsNull(rsLocEmp("BD_BNAME")) Then
            MailBody = MailBody & Space(4) & "Beneficiary Name: " & rsLocEmp("BD_BNAME") & vbCrLf
        End If
        If Not IsNull(rsLocEmp("BD_DOB")) Then
            MailBody = MailBody & Space(4) & "Beneficiary Birth Date: " & rsLocEmp("BD_DOB") & vbCrLf
        End If
        If Not IsNull(rsLocEmp("BD_RELATE")) Then
            If Len(rsLocEmp("BD_RELATE")) > 0 Then
            MailBody = MailBody & Space(4) & "Relationship: " & rsLocEmp("BD_RELATE") & vbCrLf
            End If
        End If
    End If
    'Ticket #20139 Franks 04/11/2011 - end

    frmSendEmail.txtBody.Text = MailBody
    
    MDIMain.panHelp(0).FloodType = 0
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
    MDIMain.panHelp(0).FloodType = 1


exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

Sub EmailSendingForSamuel()
    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    Dim xToEmail As String
    Dim xRAToEmail As String 'Ticket #23453 Franks 03/26/2013
    Dim xEmailSubject As String, xBranch  As String
    Dim xEmpName As String
    
    On Error GoTo ErrorHandler
    Load frmSendEmail
    'frmSendEmail.txtSubject.Text = "info:HR Termination Notice"
    'Ticket #18578
    'frmSendEmail.txtSubject.Text = "info:HR Termination Notice - " & lblEEName.Caption
    'Ticket #18755
    xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", glbTERM_Seq)
    If Len(xBranch) > 0 Then
        xBranch = xBranch & " - "
    End If
    xEmailSubject = "info:HR Termination Notice - " & xBranch & lblEEName.Caption
    frmSendEmail.txtSubject.Text = xEmailSubject
    
    MailBody = GetEmailBodyForSamuel(glbLEE_ID, glbTERM_Seq)
    MailBody = MailBody & "was terminated on " & dlpTermDate.Text & vbCrLf
    '''Ticket #23453 Franks 04/01/2013 - begin
    ''MailBody = MailBody & vbCrLf & "Managers are responsible for providing IT with a properly completed Employee Termination Form. You have 10 days to submit this form. "
    ''MailBody = MailBody & "Failure to do will result in IT closing down all system access for this individual. "
    ''MailBody = MailBody & "You will need to coordinate directly with IT if alternate arrangements are required. "
    ''MailBody = MailBody & "This form may be accessed on the Torch at this location" & vbCrLf
    ''MailBody = MailBody & "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.1.docm"
    '''Ticket #23453 Franks 04/01/2013 - end
    
    'Ticket #24475 Franks 10/22/2013 - begin - using html email
    ''Dim xLink
    ''xLink = "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.1.docm"
    ''MailBody = MailBody & "was terminated on " & dlpTermDate.Text & "<br>"
    ''MailBody = MailBody & "<b>Upon receipt of this email you must provide IT with a properly completed Employee Termination Form</b> in order to make arrangements for:<br>"
    ''MailBody = MailBody & "<ul>"
    ''MailBody = MailBody & "<li>Cellphone/blackberry"
    ''MailBody = MailBody & "<li>New emails"
    ''MailBody = MailBody & "<li>Stored email messages"
    ''MailBody = MailBody & "</ul>"
    ''MailBody = MailBody & "You have 10 days to submit this form.  Failure to do so will result in IT closing down all system access for this individual. "
    ''MailBody = MailBody & "In the event you need to maintain system privileges in this employees’ name you will need to coordinate directly with IT. "
    ''MailBody = MailBody & "Access the form here <a href='" & xLink & "'>Employee Termination Form</a><br>"
    ''MailBody = MailBody & "<br><b>In addition you must notify your 24/7 Health & Safety system administrator to remove this employee from active status.</b> "
    
    MailBody = MailBody & "was terminated on " & dlpTermDate.Text & vbCrLf
    MailBody = MailBody & "Upon receipt of this email you must provide IT with a properly completed Employee Termination Form in order to make arrangements for:" & vbCrLf & vbCrLf
    MailBody = MailBody & "  " & Chr(149) & " Cellphone/blackberry" & vbCrLf    '149 - bullet character
    MailBody = MailBody & "  " & Chr(149) & " New emails" & vbCrLf
    MailBody = MailBody & "  " & Chr(149) & " Stored email messages" & vbCrLf & vbCrLf
    MailBody = MailBody & "You have 10 days to submit this form.  Failure to do so will result in IT closing down all system access for this individual. "
    MailBody = MailBody & "In the event you need to maintain system privileges in this employees’ name you will need to coordinate directly with IT. " & vbCrLf & vbCrLf
    MailBody = MailBody & "In addition you must notify your 24/7 Health & Safety system administrator to remove this employee from active status." & vbCrLf & vbCrLf
    MailBody = MailBody & "This form may be accessed on the Torch at this location" & vbCrLf
    'MailBody = MailBody & "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.1.docm"
    MailBody = MailBody & "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.5.docm"
    'Ticket #24475 Franks 10/22/2013 - end
    
    'Ticket #24685 - Adding French to the Termination email - Begin
    xEmpName = GetEmpData(glbLEE_ID, "ED_FNAME", glbTERM_Seq) & " " & GetEmpData(glbLEE_ID, "ED_SURNAME", glbTERM_Seq)
    MailBody = MailBody & vbCrLf & vbCrLf
    MailBody = MailBody & "L'emploi de, " & xEmpName & " - No d'employé " & glbLEE_ID
    MailBody = MailBody & ", de la paie No. " & GetEmpData(glbLEE_ID, "ED_ADMINBY", glbTERM_Seq)
    MailBody = MailBody & " de la succursale de " & GetTABLDesc("EDSE", GetEmpData(glbLEE_ID, "ED_SECTION", glbTERM_Seq))
    MailBody = MailBody & " a pris fin le " & dlpTermDate.Text & ". "
    MailBody = MailBody & "Sur réception de ce courriel, vous devez fournir au département des TI le formulaire de "
    MailBody = MailBody & "cessation d'emploi " & Chr(40) & Chr(185) & Chr(41) & " dûment complété afin de prendre des dispositions pour:" & vbCrLf & vbCrLf
    MailBody = MailBody & "  " & Chr(149) & " Téléphone cellulaire / BlackBerry" & vbCrLf
    MailBody = MailBody & "  " & Chr(149) & " Nouveaux courriels" & vbCrLf
    MailBody = MailBody & "  " & Chr(149) & " Courriels enregistrés" & vbCrLf & vbCrLf
    MailBody = MailBody & "Vous avez 10 jours pour soumettre ce formulaire." & vbCrLf & vbCrLf
    MailBody = MailBody & "À défaut de faire, cela entrainera la fermeture de tous les systèmes d'accès de cette "
    MailBody = MailBody & "personne par le département des TI." & vbCrLf & vbCrLf
    MailBody = MailBody & "Dans le cas où vous avez besoin de maintenir les privilèges d'accès au système au nom de "
    MailBody = MailBody & "cet employé, vous devrez prendre des arrangements nécessaires directement avec le "
    MailBody = MailBody & "département des TI." & vbCrLf & vbCrLf & vbCrLf
    MailBody = MailBody & Chr(40) & Chr(185) & Chr(41) & " Ce formulaire peut être consulté sur le site Intranet de Samuel, en cliquant sur le lien ci-dessous." & vbCrLf & vbCrLf
    'MailBody = MailBody & "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.1.docm" & vbCrLf & vbCrLf & vbCrLf
    MailBody = MailBody & "http://torch.samuel.com/doc/it/IT%20Forms%20Library/Employee%20Termination%20Form%20v1.5.docm" & vbCrLf & vbCrLf & vbCrLf
    MailBody = MailBody & "Également, il est important d'aviser votre Administrateur en Santé Sécurité afin de changer le statut de cet employé à " & Chr(171) & " inactif " & Chr(187) & "."
    'Ticket #24685 - Adding French to the Termination email - End
    
    xToEmail = GetComPreferEmail("EMAIL_ONTERM", glbLEE_ID, glbTERM_Seq)
    If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
        xToEmail = GetComPreferEmail("EMAIL_ONTERM")
    End If
    If Len(xToEmail) > 0 Then
        'Ticket #23453 Frank 03/25/2013 - begin
        If GetComPreferEmailDetUserFlag("EMAIL_ONTERM", glbLEE_ID, glbTERM_Seq) Then
            'get Rept. Authority email address
            xRAToEmail = GetReptAuthEmail(glbLEE_ID, glbTERM_Seq)
            If Len(xRAToEmail) > 0 Then
                xToEmail = xToEmail & "; " & xRAToEmail
            End If
        End If
        'Ticket #23453 Frank 03/25/2013 - end
        frmSendEmail.txtTo.Text = xToEmail
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
    
exH:
    Exit Sub
    
ErrorHandler:
    'If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If glbWFC Then 'Ticket #21791 Franks 03/22/2012
        If Index = 3 Then 'union
            If Len(clpCode(3).Text) > 0 Then
                If Len(comDIV.Text) = 0 Then
                    comDIV.ListIndex = FindCBIndex(comDIV, lblCurDiv.Caption, 4)
                End If
            End If

            Call WFCNGSEndDateForTransferOut 'Ticket #24767 Franks 12/11/2013

        End If
    End If
End Sub

Public Sub cmdEmail_Click()
    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    Dim xToEmail As String
    Dim xAlias As String 'Ticket #27983 Franks 02/09/2016
    
    On Error GoTo ErrorHandler
    Load frmSendEmail
    If glbWFC And clpCode(1) = "TOUT" Then   'Ticket #23173 Franks 01/28/2013 - for transfer out
        frmSendEmail.txtSubject.Text = "info:HR Transfer Notice - " & lblEEName.Caption
    Else
        'frmSendEmail.txtSubject.Text = "info:HR Termination Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Termination Notice - " & lblEEName.Caption
    End If
    If glbWFC Then 'Ticket #27983 Franks 02/09/2016
        xAlias = GetEmpData(lblEENum.Caption, "ED_ALIAS")
        If Len(xAlias) > 0 Then
            frmSendEmail.txtSubject.Text = frmSendEmail.txtSubject.Text & "(" & xAlias & ")"
        End If
    End If
    
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        MailBody = GetEmailBodyForSamuel(glbLEE_ID, glbTERM_Seq)
        MailBody = MailBody & "was terminated on " & dlpTermDate.Text & vbCrLf
    Else
        If glbWFC Then 'Ticket #27983 Franks 02/09/2016
            If glbWFC And clpCode(1) = "TOUT" Then
                MailBody = "The employee below has been transferred out." & vbCrLf & vbCrLf
                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            Else
                MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            End If
            If Len(xAlias) > 0 Then
                MailBody = MailBody & "Name: " & lblEEName.Caption & "(" & xAlias & ")" & vbCrLf
            Else
                MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            End If
            'Ticket #27983 Franks 02/09/2016 - end
        Else
            MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
        End If
    End If
    If Not glbWFC Then
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
            xToEmail = GetComPreferEmail("EMAIL_ONTERM", glbLEE_ID, glbTERM_Seq)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONTERM")
            End If
        Else
            'Ticket #20317 - 'More Emails' option for everyone
            xToEmail = GetComPreferEmail("EMAIL_ONTERM", glbLEE_ID, glbTERM_Seq)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONTERM")
            End If
        End If
        frmSendEmail.txtTo.Text = xToEmail
    End If
    ' dkostka - 02/23/01 - Removed Reason from email body, added Location for WFC only.
    If glbWFC Then
        GetLocation lblEENum.Caption, LocCode, LocDesc
        MailBody = MailBody & "Location: " & LocCode & " - " & LocDesc & vbCrLf
        'Ticket #28368 Franks 05/31/2016 - "   Display RA #1's name in the box of the email.
        'MailBody = MailBody & "Reporting Authority: " & GetReportingAuthority(lblEENum.Caption) & vbCrLf
        MailBody = MailBody & lStr("Rept. Authority 1") & ": " & GetReportingAuthority(lblEENum.Caption) & vbCrLf
        If clpCode(1) = "TOUT" Then
            MailBody = MailBody & "Reason: TOUT - Transfer Out Of Unit" & vbCrLf
            If comDIV.Visible Then
                MailBody = MailBody & "Transfer To Division: " & comDIV.Text & vbCrLf
            End If
        End If
    End If
    If Not glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        MailBody = MailBody & "Date: " & dlpTermDate.Text & vbCrLf & vbCrLf
    End If
    frmSendEmail.txtBody.Text = MailBody
    
    ' dkostka - 02/23/2001 - Automated email sending for WFC.
    If glbWFC Then
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = "Sending email..."
        'Franks 05/03/04 Ticket #6105 David Hili wants to change it
        'frmSendEmail.txtTo.Text = "hotline@woodbridgegroup.com"
        frmSendEmail.txtTo.Text = glbWFCTermEmail '"termnotice@woodbridgegroup.com"
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
        MDIMain.panHelp(0).FloodType = 1
    Else
        frmSendEmail.Show 1
    End If
    
exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH
End Sub

''Private Function GetLocation(EmpNbr, ByRef LocCode As String, ByRef LocDesc As String)
''    Dim rsEmp As New ADODB.Recordset, RSTABL As New ADODB.Recordset
''
''    rsEmp.Open "SELECT ED_LOC FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
''    If rsEmp.EOF Then
''        LocCode = ""
''        LocDesc = ""
''        rsEmp.Close
''        Exit Function
''    End If
''    If Not IsNull(rsEmp("ED_LOC")) Then
''        LocCode = rsEmp("ED_LOC")
''    Else
''        LocCode = ""
''    End If
''    rsEmp.Close
''
''    RSTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='EDLC' AND TB_KEY='" & LocCode & "'", gdbAdoIhr001
''    If RSTABL.EOF Then
''        LocDesc = ""
''        RSTABL.Close
''        Exit Function
''    End If
''    LocDesc = RSTABL("TB_DESC")
''    RSTABL.Close
''End Function
''Private Function GetReportingAuthority(EmpNbr)
''    Dim rsEmp As New ADODB.Recordset, rsJobHis As New ADODB.Recordset
''    GetReportingAuthority = ""
''    rsJobHis.Open "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001
''    If Not rsJobHis.EOF Then
''        If Not IsNull(rsJobHis("JH_REPTAU")) Then
''            If IsNumeric(rsJobHis("JH_REPTAU")) Then
''                rsEmp.Open "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & rsJobHis("JH_REPTAU"), gdbAdoIhr001
''                If Not rsEmp.EOF Then
''                    GetReportingAuthority = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
''                End If
''                rsEmp.Close
''            End If
''        End If
''    End If
''    rsJobHis.Close
''End Function

Public Function GetEmpData(EmpNbr, Field As String, Optional xTERM_Seq) As String
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim retVal
    retVal = ""
    If IsMissing(xTERM_Seq) Or IsNull(xTERM_Seq) Then
        SQLQ = "SELECT " & Field & " FROM HREMP WHERE ED_EMPNBR=" & EmpNbr
    Else
        SQLQ = "SELECT " & Field & " FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr
        SQLQ = SQLQ & " AND TERM_SEQ = " & xTERM_Seq & " "
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp(Field)) Then
            retVal = Format(rsEmp(Field), "@")
        End If
    End If
    GetEmpData = retVal
    rsEmp.Close
End Function

Private Sub cmdFrankTest_Click()
    If IsWFCReptAuth(glbLEE_ID, "") Then
        glbWFC_IncePlanID = glbLEE_ID
        If glbTermTran Then
             glbWFC_IPPopFormName = "WFCEmpListWithRepTerm"
        Else
            glbWFC_IPPopFormName = "WFCEmpListWithRepTran"
        End If
        frmCheckListView.lblStDate = Date ' dlpTermDate.Text
        frmCheckListView.Show 1
    End If
End Sub

Private Sub cmdImport_Click()
    glbDocName = "Termination"
    glbDocKey = 0
    glbDocNewRecord = False
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEESTATS_Term")
    cmdImport.Visible = True
    
    If frmInAttachment.IfExist Then
        imgSec.Visible = True
        imgNoSec.Visible = False
    Else
        imgSec.Visible = False
        imgNoSec.Visible = True
    End If
    
End Sub

Private Sub cmdPhoto_Click()
Call SubPicture
End Sub

Private Sub cmdPrintSelected_Click()
Dim X%

'On Error GoTo PrntErr

If lblEEID = 0 Then
    MsgBox "        No Current Record" & Chr(10) & "Use 'FIND' to Select a Employee"
    Exit Sub
End If

If chkTermRpts(0) = True Then GoTo Prt_OK
If chkTermRpts(1) = True Then GoTo Prt_OK
If chkTermRpts(2) = True Then GoTo Prt_OK
If chkTermRpts(3) = True Then GoTo Prt_OK
If chkTermRpts(4) = True Then GoTo Prt_OK
If chkTermRpts(5) = True Then GoTo Prt_OK
'If chkTermRpts(6) = True Then GoTo Prt_OK

Exit Sub

Prt_OK:
X% = Cri_Select(1)        '0=View 1=Print
Screen.MousePointer = DEFAULT

Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
Screen.MousePointer = DEFAULT

Call RollBack '29July99 js

End Sub

Private Sub cmdPrintSelected_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub cmdTerminate_Click()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim DelayProcess As Boolean
Dim xTimminsBenefits As Boolean
Dim xPAData
Dim xPenType As String
Dim x2ndEmpID, x2ndPayComp, x1stEmpID
Dim xCurPosition

If lblEEID = 0 Then
    MsgBox "        No Current Record" & Chr(10) & "Use 'FIND' to Select a Employee"
    Exit Sub
End If

If Not chkTerms() Then Exit Sub

If glbLinamar Then 'Ticket# 8293
    Title$ = TranStr("Terminate Employee")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = "Would you rehire this employee?"
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        chkRehire.Value = False
        lblRehire.Caption = "No"
    Else
        chkRehire.Value = True
        lblRehire.Caption = "Yes"
    End If
    Msg = ""
    DoEvents
End If

If lblRptsPrinted.Visible = False Then
    If glbWFC And glbDivTranInPlant = "Y" Then 'Ticket #25384 Franks 04/22/2014
        Msg$ = ""
    Else
        Msg$ = "Not all employee reports were printed."
    End If
End If

If glbTermTran Then
    'Msg$ = Msg$ & Chr(10) & "Are you sure you want to terminate "
    Msg$ = Msg$ & Chr(10) & Chr(10) & "Are you sure you want to terminate "
Else
    Msg$ = Msg$ & Chr(10) & "Are you sure you want to transfer out "
End If
'Msg$ = Msg$ & Chr(10) & "this employee ?"
Msg$ = Msg$ & "this employee ?"
'Msg$ = Msg$ & Chr(10) & "Make sure no other info:HR Window "
'Msg$ = Msg$ & Chr(10) & "is open with this employee information showing"
DelayProcess = (CVDate(dlpTermDate) > Date) And glbLinamar
If Not DelayProcess Then
    Msg$ = Msg$ & Chr(10) & Chr(10) & Chr(10) & "Note: Make sure no other info:HR Window "
    'Msg$ = Msg$ & Chr(10) & "is open with this employee information showing."
    Msg$ = Msg$ & "is open with this employee information showing."
End If

Title$ = TranStr("Terminate Employee")
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    x2ndEmpID = FamilyDay2ndEmpNo(glbLEE_ID)
    x1stEmpID = 0
    If x2ndEmpID = 0 Then
        x1stEmpID = FamilyDay1stEmpNo(glbLEE_ID)
    End If
End If

If DelayProcess Then
    Call SaveDelayInfo
    Exit Sub
End If

If glbSamuel Then 'Ticket #20885 Franks 11/18/2011
    If glbTermTran Then 'termination only
        Call CheckReptAuth
    End If
End If

'Ticket #22682 - Release 8.0 - If the terminated employee is a reporting authority for any active employees then
'prompt the user if they want to update with another Reporting Authority otherwise clear the RA from wherever used.
xWFCPosChgEmailBody = ""
xIsWFCPosChgEmail = False
If glbTermTran Then 'termination
    If glbWFC Then 'Ticket #29343 Franks 10/25/2016
        'Call CheckWFCReptAuthExists
        'Ticket #29507 Franks 11/30/2016 don't use the function above
    Else
        Call CheckReptAuthExists
    End If
End If

If glbWFC Then 'Ticket #29507 Franks 11/30/2016
    glbWFC_CancelTransaction = False
    Call CheckWFCReptAuthExistNew
    If glbWFC_CancelTransaction Then
        Exit Sub
    End If
End If

If glbCompSerial = "S/N - 2241W" Then ' Granite Club Ticket #22056 Franks 05/22/2012
    'change the employee status to TERM first
    Call GraniteClubEmpStatusChg
End If

'Ticket #18236 - City of Timmins - Would the terminated employee be still entitled to Benefits?
If glbCompSerial = "S/N - 2375W" Then
    xTimminsBenefits = True
    Response% = MsgBox("Would this terminated employee be entitled to Benefits? ", MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2, "Entitled to Benefits?")
    If Response% = IDNO Then
        xTimminsBenefits = False
    End If
End If

If glbLinamar Then
    Dim rsWT As New ADODB.Recordset
    rsWT.Open "SELECT * FROM HR_WILL_TERM WHERE TL_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsWT.EOF Then rsWT.Delete
    rsWT.Close
End If

glbChgBenTermDate = ""
If glbWFC Then
    'Logic:
    '1. Transfer out, don't do any thing for Manulife
    '2. Termination, always terminate Benefits and Dependents
    If Not glbTermTran Then 'Transfer Out
        ''glbChgBenTermDate = dlpBenCeaseDate
        'Ticket #24451 Franks 10/15/2013
        '"   For a MLF employee, when transferring out from Canada to another location, the benefits would have to end.
        'The benefit end date should equal the Transfer Out Date
        If Len(locCertNo) > 0 Then
            If IsOutFromBenGrpMatrix(Left(comDIV.Text, 4)) Then
                If Len(dlpBenCeaseDate.Text) = 0 Then
                    glbChgBenTermDate = dlpTermDate
                Else
                    glbChgBenTermDate = dlpBenCeaseDate
                End If
            End If
        End If
    Else ' Termination
        If Len(locCertNo) > 0 And Len(dlpBenCeaseDate.Text) = 0 Then
            glbChgBenTermDate = dlpTermDate 'If this emp has Cert# and Ben End Date is blank, default it as Term Date
        Else
            glbChgBenTermDate = dlpBenCeaseDate
        End If
    End If
    
    If glbTermTran Then 'termination
        'Ticket #23575 Franks 04/12/2013 - Remove from program
        'Call WFC_PT_PenCheck   'Ticket #23117 Franks 01/28/2013
    End If
    
    xCurPosition = getEmpPostion(glbLEE_ID) 'Ticket #25911 Franks 12/17/2014
End If
MDIMain.panHelp(0).FloodType = 1

Screen.MousePointer = vbHourglass

' dkostka - 02/23/01 - Added automated email sending for WFC
If glbWFC Then
    If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
        If gsEMAIL_ONTERM Then
            If glbDivTranInPlant = "Y" Then 'Ticket #25307 Franks 04/08/2014
                '- don't send email during transfer within plants
            Else
                If WFCNonUnion(lblEENum.Caption) Then
                    ' Make sure we have needed info to send email
                    If GetEmpData(glbEmpNbr, "ED_EMAIL") = "" Then ' And Not MDIMain.mnu_File_EmailSetup.Visible Then
                        Screen.MousePointer = vbDefault
                        MsgBox GetEmpData(glbEmpNbr, "ED_FNAME") & ", please fill in your email address on the Status/Dates screen, before attempting to terminate an employee.", vbExclamation + vbOKOnly, "Missing Email Address"
                        Exit Sub
                    Else
                        If Not IsEmailSetup(glbEmpNbr) Then 'MDIMain.mnu_File_EmailSetup.Visible And Not IsEmailSetup(glbEmpNbr) Then  'lost condition afther removing menu items , should check
                            Screen.MousePointer = vbDefault
                            MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for email sending before attempting to terminate salaried employees.  Termination aborted.", vbCritical + vbOKOnly, "No Email Setup Found"
                            Exit Sub
                        End If
                    End If
                    ' Send the email
                    cmdEmail_Click
                    ' AC - dkostka - 05/03/01 - Added error checking, refuse to terminate if email didn't go through
                    If AbortTerm = True Then
                        'Screen.MousePointer = vbDefault
                        'MDIMain.panHelp(0).FloodType = 1
                        'MDIMain.panHelp(0).Caption = "Termination Aborted"
                        'MsgBox "Error sending email.  Termination aborted.", vbCritical + vbOKOnly, "Error"
                        'Exit Sub
                        'Ticket #24422 Franks 10/01/2013 - "Can't stop the termination or transfer out if the email sending doesn't work.
                        MsgBox "Error sending termination email."
                    End If
                End If
            End If
        End If
    End If
    'Ticket #16395
    'The system sends an email to pension@woodbridgegroup.com notifying them
    'that an employee has been terminated.
    If locWFCPenEligible Then 'WFCPensionEligible(lblEENum.Caption) Then
        ''If gsEMAIL_ONTERM Then
        ''    Call cmdEmailWFCPension
        ''
        ''    If AbortTerm = True Then
        ''        Screen.MousePointer = vbDefault
        ''        MDIMain.panHelp(0).FloodType = 1
        ''        MDIMain.panHelp(0).Caption = "Termination Aborted"
        ''        MsgBox "Error sending email.  Termination aborted.", vbCritical + vbOKOnly, "Error"
        ''        Exit Sub
        ''    End If
        ''End If
        
        'Pension Alert - Benficiary
        'Call WFCPensionAlerts(glbLEE_ID, Date, "Termination - " & clpCode(1).Text)
        'Call WFCPensionAlerts(glbLEE_ID, Date, "Termination ", clpCode(1).Text)
        
        'Ticket #22009 Franks 05/09/2012
        ' Create the Termination Alert and delete all other Pension Alerts.
        Call WFCPensionAlerts(glbLEE_ID, dlpTermDate.Text, "Termination - " & clpCode(1).Text, , , , "ALL")
        
        Call WFCPensionAlerts(glbLEE_ID, dlpTermDate.Text, "Termination - " & clpCode(1).Text)
        
        'Grey out Work Flow Master and Detail. Remove the Work Flow code on Termination.
        'December 4, 2009 on WFC Pension Outstanding Tasks By Dec1009.doc
        '' A Work Flow employee-based table is created to show what step the termination is in
        'If Len(clpCode(0).Text) > 0 Then
        '    Call WorkFlowUpdate(glbLEE_ID, clpCode(0).Text, dlpTermDate.Text)
        'End If
        
        'Ticket #18265 - 03/26/2010 by Frank
        If locWFCPenEarnFlag Then
            If IsNumeric(medAmount.Text) Then
                'Call Upt_PayrollTransaction(glbLEE_ID, "D", "D49", CVDate("Jan 1," & Year(dlpTermDate.Text)), dlpTermDate.Text, medAmount.Text)
                Call Upt_PayrollTransaction(glbLEE_ID, "D", "DN49", CVDate("Jan 1," & Year(dlpTermDate.Text)), dlpTermDate.Text, medAmount.Text)
            End If
        End If
    End If
    'Ticket #16748
    'gdbAdoIhr001.Execute "UPDATE HREMP SET ED_PENSION='" & Left(cboStatFlag3.Text, 1) & "' WHERE ED_EMPNBR=" & CLng(lblEEID)
    'Ticket #16616.
    'Ticket #16749 dont use Pay Group
    'If Not glbTermTran Then
    '    If Len(clpVadim2.Text) > 0 Then
    '        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VADIM2='" & Trim(clpVadim2.Text) & "' WHERE ED_EMPNBR=" & CLng(lblEEID)
    '    End If
    'End If
    
End If
'End If

If glbWFC Then 'Ticket #19266 Franks 12/02/2010
    'NGS Transactions Ticket #19266 11/25/2010 Frank
    If Not glbTermTran Then 'Transfer Out
        If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/18/2014
            'don't make any change for NGS for DivTranInPlant
        Else
            'Call WFC_NGS_Trans("Transfer Out")
            'Ticket #24184 Franks 11/12/2013
            Call WFC_NGS_Trans_TermTransferOut("Transfer Out", dlpTermDate.Text, Trim(Left(comDIV.Text, 4)), clpCode(3).Text, "")
            
            'Ticket #24767 Franks 12/11/2013
            If frmWFCBenList.Visible Then 'US NGS employees
                Call WFC_NGSBenEndDateUpt(glbLEE_ID)
            End If
        End If
    Else 'Termination
        If frmLastDay.Visible Then 'Ticket #26308 Franks 11/27/2014
            WFC_LastDayUpt (glbLEE_ID)
        End If
        If dlpDOther2.Visible Then
            'Call WFC_NGS_Trans("Termination")
            'Ticket #24184 Franks 11/12/2013
            Call WFC_NGS_Trans_TermTransferOut("Termination", dlpTermDate.Text, Trim(Left(comDIV.Text, 4)), clpCode(3).Text, dlpDOther2.Text)
        End If
        'Ticket #23948 Frank 06/24/2013
        Call WFC_UptPenDate4WithDOT(glbLEE_ID, dlpTermDate.Text)
        
        'Ticket #23247 Franks 07/22/2013
        If frmWFCBenList.Visible Then 'US NGS employees
            Call WFC_NGSBenEndDateUpt(glbLEE_ID)
        End If
    End If
Else
    'Ticket #27245 Franks 09/02/2015 non WFC Diff Benefit End Date
    If chkDiffBenEnd.Value And frmWFCBenList.Visible Then
        Call NonWFC_BenEndDateUp(glbLEE_ID) 'WFC_NGSBenEndDateUpt(glbLEE_ID)
    End If
End If

If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21518 Franks 05/03/2012
    Call AUDIT_GWL_TRANS
End If
    
rsTB.Open "Term_HRSEQ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
If rsTB.EOF And rsTB.BOF Then
    glbTERM_Seq = 1
    rsTB.AddNew
Else
    rsTB.MoveFirst 'Jaddy 10/28/99
    glbTERM_Seq = rsTB("TERM_SEQ_NEXT")
End If
rsTB("TERM_SEQ_NEXT") = glbTERM_Seq + 1
rsTB.Update

rsTB.Close
Call updStatus

''Ticket #20270 Franks 05/05/2011
'Ticket #25669 Franks 06/24/2014 - comment out the following code
'If glbTermTran Then
'    Call EEO_Process
'Else
'    'Ticket #24422 Franks 10/02/2013
'    'Transfer Out should not delete the EEO records.-
'End If


If Not AUDITTERM() Then MsgBox "ERROR - AUDIT FILE"

'Ticket #23409 - Samuel - Add Discipline Audit
If glbSamuel Then
    If Not AUDITCOUNSEL() Then MsgBox "ERROR - AUDIT FILE"
End If

If glbWFC Then 'Ticket #13448
    'Call AUDIT_MANULIFE_TRANS
    'Ticket #24184 Franks 11/12/2013
    If glbDivTranInPlant = "Y" Then
        'Ticket #25248 Franks 03/24/2014
        '"   Don't create the MLF audit record. A change within a plant doesn't need to be exported to MLF.
    Else
        Call AUDIT_MANULIFE_TRANS_TermTransferOut(locCertNo)
    End If
Else
    'Ticket #18668 - Update Audit Table with the Benefit End Date if entered.
    If IsDate(dlpBenCeaseDate) Then
        If Not AUDIT_BenefitEndDate() Then MsgBox "ERROR - AUDIT FILE - Benefit End Date"
    End If
End If

'Ticket #16395 - Pension System
If glbWFC Then
    If glbTermTran Then
        'Termination -
        toSOURCE = "IHR Termination" 'Ticket #19954
        xPAData = "PA"
        If locWFCPenEarnFlag Then
            If IsNumeric(medAmount.Text) Then
                xPAData = "PA|" & Trim(Str(medAmount.Text))
            End If
        End If
        Call WFCPensionMasUpt(glbLEE_ID, "Termination", dlpTermDate.Text, clpCode(1).Text, Year(CVDate(dlpTermDate)), xPAData)
        If clpCode(1).Text = "DECD" Then
            'One employee can have one DBS plus other DB pensions, such as DBKIPL
            'Employee Dan Dubblestyne had DBS and DBKIPL pensions, create other pensions for this year with status "D"
            
            'Ticket #26707 Franks 02/25/2015 - begin
            'xPenType = getDBType(locSection, locUnion, "PenType")
            xPenType = getDBType(locSection, locUnion, "PenType", GetEmpData(glbLEE_ID, "ED_DOH"))
            'Ticket #26707 Franks 02/25/2015 - end
            Call WFCOtherPenUpt(glbLEE_ID, glbSIN, Year(dlpTermDate.Text), "", xPenType, "D", dlpTermDate.Text, dlpTermDate.Text, "DB")
        End If
        
        'Ticket #22009 Franks 05/09/2012
        'delete other Alerts which were created in Termination
        Call WFCPensionAlerts(glbLEE_ID, dlpTermDate.Text, "Termination - " & clpCode(1).Text, , , , "Y")
        
        If xIsWFCPosChgEmail Then  'Ticket #29343 Franks 10/25/2016
            If gsEMAIL_ONPOSITION Then
                If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
                    Call WFCPubPosChangedcmdEmail(glbLEE_ID, xWFCPosChgEmailBody, "info:HR Position Reporting Authority Change Notice")
                End If
            End If
        End If
    Else
        'Transfer Out - Trans Date, New Div
        toSOURCE = "IHR Transfer Out" 'Ticket #19954
        Call WFCPensionMasUpt(glbLEE_ID, "Transfer Out", dlpTermDate.Text, Left(comDIV.Text, 4), Year(CVDate(dlpTermDate)))
        'Ticket #19678 Franks 01/24/2011
        'On Transfer out:   (Plant Code equals "TILB" and Union Code is "C127") or (Plant Code equals "WHBY" and Union Code is "C222") the Hire Code equals "N".
        Call WFCHireCode(glbLEE_ID)
        
        'Ticket #21677 Franks 03/14/2012
        If clpCode(3).Text = "NONE" Or clpCode(3).Text = "EXEC" Then
            'Call locWFCUpdPAMaster(glbLEE_ID, dlpTermDate.Text, locSIN)
            'Ticket #24184 Franks 11/12/2013
            Call locWFCUpdPAMaster_TermTransferOut(glbLEE_ID, dlpTermDate.Text, locSIN, Left(lblCurDiv.Caption, 4), locPayrollID)
        End If
        
    End If
    
    If gsEMAIL_ONTERM Then
        If locWFCPenEligible Then
            Call cmdEmailWFCPension
            
            If AbortTerm = True Then
                'Screen.MousePointer = vbDefault
                'MDIMain.panHelp(0).FloodType = 1
                'MDIMain.panHelp(0).Caption = "Termination Aborted"
                'MsgBox "Error sending email.  Termination aborted.", vbCritical + vbOKOnly, "Error"
                'Exit Sub
                'Ticket #24422 Franks 10/01/2013 - "Can't stop the termination or transfer out if the email sending doesn't work.
                'MsgBox "Error sending Pension email."
            End If
        End If
    End If
    
End If

Call UpdPositionCCAC

If glbLambton Then
    Call UpdPositionMulti
Else
    '8.0 - Ticket #22682 - Update Current Positions with End Date.
    If glbTermTran And chkPosEndDate Then
        Call UpdPositionEndDate
    End If
End If

Call UpdPaymentTypeVadim

If glbCompSerial = "S/N - 2394W" Then  'St. John #14752
    Call UpdEmpType("X")
End If

If glbCompSerial = "S/N - 2370W" Then  'David Chapman's Ice Cream Limited - Ticket #15601
    Call UpdEmpStatus("T")
End If

'Ticket #18236 - City of Timmins - Clear the Benefit Group as the employee will no longer be getting the benefit.
If glbCompSerial = "S/N - 2375W" Then
    If xTimminsBenefits = False Then
        Call ClearEmpBenefitGroup
    End If
End If

If Not modTermMove() Then Exit Sub

EID& = CLng(lblEEID)
TermDate$ = dlpTermDate.Text

'If Not Term_Superv() Then Exit Sub  'laura
'Ticket #24184 Franks 11/12/2013
If Not Term_Superv_General(glbLEE_ID) Then Exit Sub  '

'If Not Term_Reviewer() Then Exit Sub  'George Apr 4,2006 #10595
'Ticket #24184 Franks 11/12/2013
If Not Term_Reviewer_General(glbLEE_ID) Then Exit Sub

MDIMain.panHelp(0).FloodPercent = 100
If Not glbTermTran Then
    Call updFollow
    If glbSQL Or glbOracle Then
        'Ticket #24552 Franks 11/01/2013 - don't change the emp no here,  will change it in transfer in if the emp no was changed
        'gdbAdoIhr001.Execute "update HR_PHOTO set pt_empnbr=" & fglbEMPNBR & " where pt_empnbr=" & EID&
    End If
Else
    If glbSQL Or glbOracle Then
        'Ticket #20367 - Jerry said we should not delete the Photo and also should rehire and
        'we should be able to see Photo on Demographics screen of the Terminated employee
        'If Not glbLinamar Then  'Ticket #13799 - Do not delete Employee Photo for Linamar
        '    gdbAdoIhr001.Execute "delete from HR_PHOTO where PT_EMPNBR=" & EID&
        'End If
    End If
End If


Call UpdEHScorrective(glbLEE_ID, lblEEName.Caption)

'Call NukeEE2(EID&)
'Ticket #24184 Franks 11/12/2013
Call NukeEE2_General(EID&)

'Ticket #25911 Franks 12/17/2014 - begin
'Release 8.1 - update the Budgeted Position
If glbWFC Then
    Call mod_Upd_Pos_Budget_WFC(xCurPosition, "")
End If
'Ticket #25911 Franks 12/17/2014 - end
    
MDIMain.panHelp(0).FloodPercent = 0

lblEEID = 0
'~~~~~~~~~~~~~~~~~~~~~~~~~'added by RAUBREY 5/23/97 ~~~~~~~~~~~~~~~~~~~~~~
rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #27829
    rsT_PARCO("PC_NUMBER_EMPLOYEES") = modECount_FamilyDay
Else
    rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
End If
rsT_PARCO.Update
rsT_PARCO.Close
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Auto Pop-up Email Notification for Termination
If gsEMAIL_ONTERM Then
    If Not glbWFC Then
        If Not IsEmailSetup(glbEmpNbr) Then
            MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for email sending. ", vbCritical + vbOKOnly, "No Email Setup Found"
        Else
            Screen.MousePointer = DEFAULT
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352
                If glbTermTran Then 'Ticket #23453 Franks 04/01/2013 - Termination only
                    Call EmailSendingForSamuel
                End If
            Else
                Call cmdEmail_Click
                Unload frmSendEmail
            End If
            Screen.MousePointer = HOURGLASS
        End If
    End If
End If

If Not glbTermTran Then 'Transfer Out
    SQLQ = "UPDATE Term_HREMP SET ED_OMERS=" & Date_SQL(dlpTermDate.Text)
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    gdbAdoIhr001X.Execute SQLQ
    Call UptLUserLDateLTime(glbTERM_Seq) 'Ticket #24355 Franks 09/17/2013
End If

If glbAdv Then 'Ticket #15074
    Call Employee_Master_Integration(glbLEE_ID, "T" & Trim(Str(glbTERM_Seq)))
Else
    Call Employee_Master_Integration(glbLEE_ID, , , glbTERM_Seq)
End If

lblEEName = "Last Employee Reviewed was Terminated"
glbLEE_ID = 0
lblRptsPrinted.Visible = False
dlpTermDate.Text = ""
clpCode(1).Text = ""

'Ticket #21238 - County of Oxford
If glbCompSerial = "S/N - 2259W" Then
    chkRehire.Value = False
    lblRehire.Caption = "No"
Else
    chkRehire.Value = True
End If

txtComments.Text = ""
'cboStatFlag3.Text = "" 'Ticket #16748
' danielk - 12/31/2002 - Ticket #2524 - 7.0 Priority C changes, commented out next line
'cmdClose.SetFocus
Unload frmEEBASIC 'Add by Frank Aug 17, 01

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    If x2ndEmpID > 0 Then
        SQLQ = "UPDATE HREMP SET ED_USER_TEXT2 = NULL WHERE ED_EMPNBR = " & x2ndEmpID
        gdbAdoIhr001.Execute SQLQ
        
        Title$ = TranStr("Terminate Employee")
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        x2ndPayComp = GetEmpData(x2ndEmpID, "ED_SALDIST")
        Msg = glbLEE_FName & " " & glbLEE_SName & " has another Employee Master record for Payroll " & x2ndPayComp & " "
        Msg = Msg & Chr(10) & Chr(10) & "Should this record be terminated as well?"
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        
        If Not Response% = IDNO Then    ' Evaluate response
            Call FamilyDay2ndTermScreen(x2ndEmpID)
            Exit Sub
        End If
    Else
        If x1stEmpID > 0 Then
            SQLQ = "UPDATE TERM_HREMP SET ED_USER_TEXT2 = NULL WHERE TERM_SEQ = " & glbTERM_Seq
            gdbAdoIhr001.Execute SQLQ
            SQLQ = "UPDATE HREMP SET ED_BADGEID = NULL WHERE ED_EMPNBR = " & x1stEmpID
            gdbAdoIhr001.Execute SQLQ
        
            Title$ = TranStr("Terminate Employee")
            DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
            x2ndPayComp = GetEmpData(x1stEmpID, "ED_SALDIST")
            Msg = glbLEE_FName & " " & glbLEE_SName & " has another Employee Master record for Payroll " & x2ndPayComp & " "
            Msg = Msg & Chr(10) & Chr(10) & "Should this record be terminated as well?"
            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
            
            If Not Response% = IDNO Then    ' Evaluate response
                Call FamilyDay2ndTermScreen(x1stEmpID)
                Exit Sub
            End If
        End If
    End If
End If


Call UnloadFrms

If glbWFC Then 'Ticket #25221 Franks 03/18/2014
    If glbDivTranInPlant = "Y" Then
        glbTran_Seq = glbTERM_Seq
        glbCandidate = 0
        glbHRSoftType = ""
        Load frmETRANIN
        frmETRANIN.ZOrder 0
    Else
        glbTERM_Seq = 0
    End If
Else
    glbTERM_Seq = 0 'Ticket #24729 Franks 01/24/2014
End If

'Ticket #24767 Franks 12/11/2013 - close this form for both termination and transfer out
'If glbTermTran Then 'Ticket #23247 Franks 07/23/2013
    Unload Me 'close this form to keep as same as Enter a Leave
'End If

MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdTerminate_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub


Private Function CompTime()
Dim rsTD As New ADODB.Recordset
Dim rsTE As New ADODB.Recordset
Dim xCT, xOT, xEmpnbr, SavEmp, SQLQ, xlen

CompTime = False

gdbAdoIhr001.Execute "DELETE  FROM HRENTWRK " & in_SQL(glbIHRDBW) & " WHERE TE_WRKEMP='" & glbUserID & "'"

xEmpnbr = lblEEID

rsTD.Open "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenKeyset

If rsTD.EOF Then GoTo XExit

rsTE.Open "HRENTWRK", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdTableDirect

xOT = 0
xCT = 0


Do While Not rsTD.EOF
    If Val(rsTD("AD_EMPNBR")) = Val(xEmpnbr) Then
        If Left(rsTD("AD_REASON"), 2) = "CT" Then xCT = xCT + rsTD("AD_HRS")
        If Left(rsTD("AD_REASON"), 2) = "OT" Then xOT = xOT + rsTD("AD_HRS")
    End If
    rsTD.MoveNext
Loop
'If xCT <> 0 Or xOT <> 0 Then   'Hemu - Commented this becuse if there are not CT and/or OT records
                                'then no entry is made into HRENTWRK table and so even the Vacation & Sick record is not displayed
  rsTE.AddNew
  rsTE("TE_COMPNO") = "001"
  rsTE("TE_EMPNBR") = xEmpnbr
  rsTE("TE_REASON_TABL") = "ADRE"
  rsTE("TE_REASON") = "CTOT"
  rsTE("TE_EARNHRS") = xOT
  rsTE("TE_USEDHRS") = xCT
  rsTE("TE_WRKEMP") = glbUserID
  rsTE.Update
'End If

rsTD.Close
rsTE.Close
XExit:
CompTime = True

End Function


Private Sub Cri_EE()
Dim EECri As String

If Len(lblEEID) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} = " & Val(lblEEID) & " "
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

Private Function Cri_Select(RptTo As Integer)
Dim X%
Dim strWHand As String
Dim EECri As String
Dim glbstrSelCri1 As String

On Error GoTo CRW_Err

If RptTo = 1 Then
    If Not PrtForm(TranStr("Termination Reports"), Me) Then Exit Function
End If

Screen.MousePointer = HOURGLASS

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

glbiOneWhere = False
glbstrSelCri = ""

If Len(lblEEID) = 0 Then Exit Function
glbstrSelCri = "{HREMP.ED_EMPNBR} = " & Val(lblEEID) & " "
glbiOneWhere = True
' reports names

If chkTermRpts(0).Value = True Then
    Call SELATTWRK
    glbstrSelCri1 = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhs.rpt"
    Me.vbxCrystal.WindowTitle = "Attendance History Report - Summarized"
    Me.vbxCrystal.Formulas(0) = TranStr("descGroup1 = 'Termination'")
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GH2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(3) = "GF2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 5 + 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
    End If
    Me.vbxCrystal.Formulas(0) = TranStr("DESCGROUP1 = 'Termination :'")
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
    Me.vbxCrystal.Formulas(2) = "DATERANGE = ''"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"

    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If chkTermRpts(1).Value = True Then
    If chkTermRpts(0).Value = False Then Call SELATTWRK
    glbstrSelCri1 = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhd.rpt"
    Me.vbxCrystal.WindowTitle = "Attendance History Report - Detail"
    Me.vbxCrystal.Formulas(0) = TranStr("descGroup1 = 'Termination'")
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GH2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(3) = "GF2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 5 + 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
    End If
    Me.vbxCrystal.Formulas(0) = TranStr("DESCGROUP1 = 'Termination :'")
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
    Me.vbxCrystal.Formulas(2) = "DATERANGE = ''"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"

    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If chkTermRpts(2).Value = True Then
    glbstrSelCri1 = glbstrSelCri & " AND {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
    Call EmpWrk
    Call setRptLabel(Me, 1)
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        'Changed by Frank Apr 5,2002 for the 20533 error, "cannot open database"
        'If the Databases are not in as same folder as reports are
        'For For X% = 1 To 9 '7
        For X% = 1 To 12
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        'Me.vbxCrystal.Password = gstrAccPWord$
        'Me.vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Formulas(51) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
    Me.vbxCrystal.Formulas(52) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
    Me.vbxCrystal.Formulas(53) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
    Me.vbxCrystal.Formulas(54) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "
    
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprofil.rpt"
    Me.vbxCrystal.WindowTitle = "Employee Profile Report"
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
'    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If chkTermRpts(3).Value = True Then
    glbstrSelCri1 = glbstrSelCri & " AND {HRENTWRK.TE_WRKEMP}='" & glbUserID & "'"
    X% = CompTime()
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5.rpt"
    Me.vbxCrystal.WindowTitle = "Entitlements Report"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    'Me.vbxCrystal.SectionFormat(0) = "GH1;F;X;X;X;X;X;X"
    'Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 6
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        Me.vbxCrystal.DataFiles(7) = glbIHRDBW
        ' set security for database
'        vbxCrystal.Password = gstrAccPWord$
'        vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If chkTermRpts(4).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzfollup.rpt"
    Me.vbxCrystal.WindowTitle = lStr("Follow-ups Report")
    Me.vbxCrystal.Formulas(0) = TranStr("descGroup1 = 'Termination'")
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    'Changed by Frank Apr 5,2002 for the 20533 error, "cannot open database"
    'If the Databases are not in as same folder as reports are
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    For X% = 0 To 5
    '        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    '    Next X%
    'End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If chkTermRpts(5).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzmastr1.rpt"
    Me.vbxCrystal.WindowTitle = "Employee Comments"
    Me.vbxCrystal.Formulas(0) = TranStr("descGroup1 = 'Termination'")
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    'Changed by Frank Apr 5,2002 for the 20533 error, "cannot open database"
    'If the Databases are not in as same folder as reports are
    'uncommented by Bryan 05/Dec/05 Ticket#9907
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 12
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
    End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

lblRptsPrinted.Visible = True

Exit Function

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

Resume Next

End Function
Sub load_SamuelPLANT_List()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim rsDiv As New ADODB.Recordset
On Error GoTo Div_List_Err

SQLQ = "SELECT DISTINCT HRTABL.* FROM HRTABL WHERE TB_NAME = 'EDAB' "
SQLQ = SQLQ & "ORDER BY TB_DESC "
rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic


If Not rsDiv.EOF Then
    comDIV.Clear
    Do Until rsDiv.EOF
        comDIV.AddItem Left(rsDiv("TB_KEY") & "   ", 4) & " - " & rsDiv("TB_DESC")
        rsDiv.MoveNext
    Loop
End If

Exit Sub

Div_List_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Depts", "HRDept", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub
Sub load_PLANT_List()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim rsDiv As New ADODB.Recordset
On Error GoTo Div_List_Err

'SQLQ = "Select DISTINCT HR_DIVISION.* FROM HR_DIVISION "
'SQLQ = SQLQ & " ORDER BY [DIV] "
SQLQ = "SELECT DISTINCT HRTABL.* FROM HRTABL WHERE TB_NAME = 'EDSE' "
SQLQ = SQLQ & "ORDER BY TB_DESC "
rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic


If Not rsDiv.EOF Then
    comDIV.Clear
    Do Until rsDiv.EOF
        comDIV.AddItem Left(rsDiv("TB_KEY") & "   ", 4) & " - " & rsDiv("TB_DESC")
        rsDiv.MoveNext
    Loop
End If

Exit Sub

Div_List_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Depts", "HRDept", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub
Sub load_DIV_List()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim rsDiv As New ADODB.Recordset
On Error GoTo Div_List_Err

SQLQ = "Select DISTINCT HR_DIVISION.* FROM HR_DIVISION "
If glbWFC Then
    'Ticket #21677 Franks 03/13/2012
    SQLQ = SQLQ & "WHERE (1=1) " 'for all Divisions, they can do Union transfer in the same Div
    If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/18/2014
        SQLQ = SQLQ & "AND DV_SECTION IN (SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID & ")"
        SQLQ = SQLQ & "AND DIV NOT IN (SELECT ED_DIV FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID & ")"
    End If
Else
    SQLQ = SQLQ & "WHERE DIV NOT IN (SELECT ED_DIV FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID & ")"
End If
SQLQ = SQLQ & " ORDER BY [DIV] "

rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic


If Not rsDiv.EOF Then
    comDIV.Clear
    Do Until rsDiv.EOF
        If rsDiv("DIV") <> "ALL" Then
            If glbWFC Then 'Ticket #30472 Franks 08/11/2017
                If Left(rsDiv("DIVISION_NAME"), 2) = "Z " Then
                    '"   Inactive Divisions should not show in the list.
                    'Debug.Print rsDiv("DIV") & " " & rsDiv("DIVISION_NAME")
                Else
                    comDIV.AddItem rsDiv("DIV") & " - " & rsDiv("DIVISION_NAME")
                End If
            Else
                comDIV.AddItem rsDiv("DIV") & " - " & rsDiv("DIVISION_NAME")
            End If
        End If
        rsDiv.MoveNext
    Loop
End If

Exit Sub

Div_List_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Depts", "HRDept", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

Private Sub comDIV_Click()
Dim SQLQ, xDiv
lblTitle(4).Visible = False
lblEMPNo.Visible = False
fglbEMPNBR = 0
xDiv = Left(comDIV, 3)
If Len(xDiv) > 0 And IsNumeric(xDiv) Then
    If glbLinamar Then
        Call setNewEmpNbr
    Else 'wfc
        'Ticket #15537, keep the same employee #
        'Call setNewEmpNbrALL
        fglbEMPNBR = glbLEE_ID
    End If
    lblTitle(4).Visible = True
    lblEMPNo.Visible = True
    lblEMPNo.Caption = ShowEmpnbr(fglbEMPNBR)
End If

If glbWFC Then 'Ticket #24767 Franks 12/11/2013
    Call WFCNGSEndDateForTransferOut
End If

End Sub

Private Sub comDIV_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpDOther2_LostFocus()
Call UptData2fromDOT
End Sub

Private Sub dlpPosEndDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpTermDate_Change()
    Call WFCDeemedPEsetup(dlpTermDate.Text)
End Sub

Private Sub WFCDeemedPEsetup(xDATE)
    If glbWFC Then 'Ticket #18265
        locWFCPenEarnFlag = False
        If glbTermTran Then
            If locWFCPenEligible Then
                If IsDate(xDATE) Then
                    locWFCPenEarnFlag = IfWFCPenEarn(glbLEE_ID, xDATE)
                End If
                Call setPenEarningsBox(locWFCPenEarnFlag)
            End If
        End If
    End If
End Sub

Function IfWFCPenEarn(EmpNbr, xTermDate) As Boolean
    Dim rsEmp As New ADODB.Recordset
    Dim rsTABL As New ADODB.Recordset
    Dim SQLQ As String
    Dim retVal As Boolean
    Dim tFlag As Boolean
    Dim xValidDate As Boolean 'Ticket #26489 Franks 01/07/2015
    'Ticket #18265
    'If Eligible for Pension = Yes and Union Code = NON and if there is no
    'Payroll Transaction  file for the current year loaded for Payroll Code E01,
    'display a line below "Benefit End Date" saying "Pensionable Earnings".
    'They must enter a value
    
    'More change for this:
    'Pensionable Earnings should say "Deemed PE". Only want to display this if the DOT's year is in the current year.
    'see Pension Tests - April0810.docx
    retVal = False
    tFlag = False
    If IsDate(xTermDate) Then 'dlpTermDate.Text) Then
        'If Year(dlpTermDate.Text) = Year(Date) Then
        'Ticket #26489 Franks 01/07/2015 - begin
        'If Year(xTermDate) = Year(Date) Then
        'Suggestion: Keep the same logic for Deemed PE but add a condition saying that if the month of the termination date is December and the year of the termination date is last year,
        'display the Deemed PE. If the termination month is before December, then the Deemed PE isn’t displayed. - Jerry
        xValidDate = False
        If Year(xTermDate) = Year(Date) Then xValidDate = True
        If Year(xTermDate) + 1 = Year(Date) Then
            If month(xTermDate) = 12 Then
                xValidDate = True
            End If
        End If
        'Ticket #26489 Franks 01/07/2015 - end
        If xValidDate Then
            rsEmp.Open "SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenStatic
            If Not rsEmp.EOF Then
                If Not IsNull(rsEmp("ED_ORG")) Then
                    If rsEmp("ED_ORG") = "NONE" Then
                        tFlag = True
                    End If
                End If
            End If
            rsEmp.Close
            If tFlag Then
                If IsDate(xTermDate) Then
                    SQLQ = "SELECT * FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR = " & EmpNbr & " "
                    SQLQ = SQLQ & "AND YEAR(PT_PAYSTART) = " & Year(xTermDate) & " "
                    SQLQ = SQLQ & "AND PT_PAYCODE = 'EN01' " '"AND PT_PAYCODE = 'E01' "
                    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If rsTABL.EOF Then
                        retVal = True
                    End If
                    rsTABL.Close
                End If
            End If
        End If
    End If
    IfWFCPenEarn = retVal

End Function

Private Sub setPenEarningsBox(xFlag)
    lblTitle(9).Top = 1270
    medAmount.Top = 1270
    lblTitle(9).Visible = xFlag
    medAmount.Visible = xFlag
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMETERM"
    fglbNew = False
    
    Call SET_UP_MODE
    
    'Ticket #21543 Franks 02/07/2012 No Tab Order before
    'Ticket #21994 - User are getting error when they do not have Termination security - gSec_Upd_Terminations
    If Panel3D1.Enabled = True Then
        dlpTermDate.SetFocus
    End If
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMETERM"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMETERM"
glbchkSum = False  'Jaddy 11/9/99

Screen.MousePointer = HOURGLASS
' dkostka - 02/23/01 - Email sending not coded for SQL, WFC works automatically.  Hide the button in these cases.
'jaddy turned this button on for sql and oracle. Hemu will test this. 5/25/2005

If glbWFC Then
    Call WFCScreenSetup 'Ticket #26308 Franks 11/27/2014
End If



If glbSamuel Then 'Ticket #20884 Franks 10/20/2011
    lblTitle(3).Caption = "To " & lStr("Administered By") '"PLANT"
End If

If glbTermTran Then
    lblTitle(5).Top = 960
    dlpBenCeaseDate.Top = 960
    If Not glbWFC Then 'Ticket #27245 Franks 09/01/2015 non WFC Diff Benefit End Date
        fraDiffBenEnd.Left = lblTitle(5).Left
        fraDiffBenEnd.Top = dlpBenCeaseDate.Top + 290
        fraDiffBenEnd.BorderStyle = 0
        fraDiffBenEnd.Visible = True
    End If
End If

panTermRpts.Caption = TranStr(panTermRpts.Caption)
lblTitle(0).Caption = TranStr(lblTitle(0).Caption)
dlpTermDate.Tag = TranStr(dlpTermDate.Tag)
lblTitle(1).Caption = TranStr(lblTitle(1).Caption)
clpCode(1).Tag = TranStr(clpCode(1).Tag)

cmdTerminate.Caption = TranStr(cmdTerminate.Caption)
cmdTerminate.Tag = TranStr(cmdTerminate.Tag)

Screen.MousePointer = DEFAULT

If glbLEE_ID = 0 And (Not glbtermopen) Then frmEEFIND.Show 1
If glbLEE_ID > 0 Then
    If glbWFC And glbCandidate > 0 Then 'Ticket #24184 Franks 10/28/2013
        'donot show the form
    Else
        Me.Show
    End If
    Call cll_EEFind(Me)
Else
    Unload Me
    Exit Sub
End If

'Grey out Work Flow Master and Detail. Remove the Work Flow code on Termination.
'December 4, 2009 on WFC Pension Outstanding Tasks By Dec1009.doc
'If glbWFC Then
'    'Ticket #16395 Type of Work Flow - Begin
'    If glbTermTran Then
'        locWFCPenEligible = WFCPensionEligible(glbLEE_ID)
'        lblTitle(8).Top = 1270
'        clpCode(0).Top = 1270
'        lblTitle(8).Visible = True
'        clpCode(0).Visible = True
'        If locWFCPenEligible Then
'            lblTitle(8).FontBold = True
'        End If
'    End If
'    'Ticket #16395 Type of Work Flow - End
'End If
'If glbWFC Then
'    If glbTermTran Then
'        locWFCPenEligible = WFCPensionEligible(glbLEE_ID)
'    End If
'End If

If Not glbTermTran Then 'Transfer Out
    lblImport.Caption = "Transfer Out" 'Ticket #27983 Franks 02/09/2016
    
    If glbSamuel Then 'Ticket #20884 Franks 10/20/2011
        Call load_SamuelPLANT_List
    Else
        Call load_DIV_List
    End If
    'If glbWFC Then
    '    Call load_PLANT_List
    'End If
    clpCode(1) = "TOUT"
    clpCode(1).Enabled = False
    lblTitle(3).Visible = True
    comDIV.Visible = True
    
    'Ticket #21677 Franks 03/13/2012 - add Union
    If glbWFC Then
        If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/17/2014
            Call WFCDivTranSamePlantScreen
        Else
            Call WFCNormalTranOutScreen
        End If
    End If

    '8.0 - Ticket #22682 - Option to update Current Positions with End Date.
    chkPosEndDate.Visible = False
    dlpPosEndDate.Visible = False
End If

If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    Me.Left = 0
End If
Me.WindowState = vbMaximized

Call LoadStatusFlag

lblRptsPrinted.Visible = False
Screen.MousePointer = HOURGLASS
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call PhotoFormLoad

MDIMain.panHelp(0).Caption = TranStr("Proceed with Termination ")

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

If Not gSec_Upd_Terminations Then
    chkRehire.Enabled = False
    chkSum.Enabled = False
    chkTermRpts(0).Enabled = False
    chkTermRpts(2).Enabled = False
    chkTermRpts(3).Enabled = False
    chkTermRpts(4).Enabled = False
    chkTermRpts(5).Enabled = False
'    chkTermRpts(6).Enabled = False
    cmdTerminate.Enabled = False
    Panel3D1.Enabled = False
    panTermRpts.Enabled = False
   ' clpCode(1).Enabled = False
    txtComments.Enabled = False
  '  dlpTermDate.Enabled = False
    
End If

If Not gSec_Summarize_Attendance And glbLinamar Then chkSum.Enabled = False

'Ticket #21238 - County of Oxford
If glbCompSerial = "S/N - 2259W" Then
    chkRehire.Value = False
    lblRehire.Caption = "No"
End If

glbDocName = "Termination"
If gsAttachment_DB Then
    lblImport.Visible = True
    cmdImport.Visible = True
    glbDocKey = 0
    If glbWFC And glbDivTranInPlant = "Y" Then  'Ticket #25221 Franks 03/17/2014
        'don't show attachment
        lblImport.Visible = False
        cmdImport.Visible = False
    Else
        If frmInAttachment.IfExist Then
            imgSec.Visible = True
            imgNoSec.Visible = False
        Else
            imgSec.Visible = False
            imgNoSec.Visible = True
        End If
    End If
End If

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Left = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)


    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmETERM = Nothing
End Sub

Private Function InputHREMPEQU_DOT(EmpN As Long)
Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_EMPNBR = "
SQLQ = SQLQ & EmpN


dynEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmp.RecordCount > 0 Then
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    'SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = " & Date_SQL(dlpTermDate.Text) & " "
    SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = " & Date_SQL(dlpTermDate.Text) & ", EQ_TYPE = 'T' "
    SQLQ = SQLQ & "WHERE HREMPEQU.EQ_EMPNBR = " & EmpN
    gdbAdoIhr001.Execute SQLQ
End If

End Function

Private Function modTermMove()
'Ticket #24184 Franks 11/12/2013
'Note for developers: there is another funcion 'modTermMoveWFCTransferOut' which should do the same function
'                     any change here should make the same change in modTermMoveWFCTransferOut, this is for WFC Transfer Out
Dim X%
Dim EEID&, TReason$, DtTm  As Variant, TRDesc$
Dim TComment$
Dim TRehire$
Dim TCause

Screen.MousePointer = HOURGLASS

modTermMove = False
DtTm = dlpTermDate.Text
EEID& = lblEEID
TReason$ = clpCode(1).Text
TComment$ = txtComments
TRehire$ = lblRehire
TRDesc$ = clpCode(1).Caption
TCause = clpCode(2).Text

gdbAdoIhr001.BeginTrans
'gdbAdoIhr001X.BeginTrans

X% = TERM_LIST(EEID&, DtTm, TReason$, TRDesc$, TComment$, TRehire$, TCause)
MDIMain.panHelp(0).FloodPercent = 5
X% = TERM_BASIC(EEID&)
MDIMain.panHelp(0).FloodPercent = 10
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EDUCSEM(EEID&)                  'laura nov 5, 1997
MDIMain.panHelp(0).FloodPercent = 13      '
If Not X Then GoTo modTermMoveErr_Msg    '
X% = TERM_ATTENDANCE(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 15
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_ATTENDANCE_HISTORY(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 18
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_JOB(EEID&)
MDIMain.panHelp(0).FloodPercent = 20
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_PERFORM(EEID&)
MDIMain.panHelp(0).FloodPercent = 22
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_SALARY(EEID&)
MDIMain.panHelp(0).FloodPercent = 25
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HealthSafety(EEID&)
MDIMain.panHelp(0).FloodPercent = 28
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_BENEFITS(EEID&)
MDIMain.panHelp(0).FloodPercent = 30
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DEPEND(EEID&)
MDIMain.panHelp(0).FloodPercent = 31
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HealthCost(EEID&)
MDIMain.panHelp(0).FloodPercent = 32
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_OHS_Contact(EEID&)
MDIMain.panHelp(0).FloodPercent = 35
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_COMMENTS(EEID&)               'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 38
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_COBRA(EEID&)
MDIMain.panHelp(0).FloodPercent = 39
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_OHS_Corrective(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_ROOT_CAUSES(EEID&)
If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
    MDIMain.panHelp(0).FloodPercent = 40
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HealthReoccurrence(EEID&)
End If

If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_CLAIM_MEDICAL(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_FORM7_SECTIONS(EEID&)

'Ticket #21463
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_FORM9(EEID&)

MDIMain.panHelp(0).FloodPercent = 43
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DOLENT(EEID&)

'Ticket #28789 - Actual Amounts Details
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DOLENT_ACTDTL(EEID&)

MDIMain.panHelp(0).FloodPercent = 45
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_ENTHRS(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 46
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EARN(EEID&)                   'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 48
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EDU(EEID&)                    'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 50
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPSKL(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 52
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_TRADE(EEID&)                  'FRANK 4/5/2000
If Not X Then GoTo modTermMoveErr_Msg
MDIMain.panHelp(0).FloodPercent = 53
X% = TERM_COUNSEL(EEID&)                ' dkostka - 10/02/2001
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HREMPHIS(EEID&)                ' Hemu - 06/30/2004
If Not X Then GoTo modTermMoveErr_Msg
'If glbWFC Then
X% = TERM_EMPOTHER(EEID&)                  'FRANK 11/05/2004
If Not X Then GoTo modTermMoveErr_Msg
'End If
X% = TERM_USERDEFINE_TABLE(EEID&)          'Hemu - 02/28/2008
If Not X Then GoTo modTermMoveErr_Msg

X% = TERM_SUCCESSION(EEID&)          'George 04/04/2006 #10595
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_LANGUAGE(EEID&)          'George 04/04/2006 #10595
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMP_FLAGS(EEID&)          'Bryan 05/04/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_GLDIST(EEID&)             'Bryan 05/04/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPADP(EEID&)                  'FRANK 06/08/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPPAYROLL_TRANSACTION(EEID&)  'FRANK 03/18/2010 Ticket #18232
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_FOLLOW_UP(EEID&)  'Hemu 08/27/2010 Ticket #18668
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HREEO(EEID&)  'Ticket #25669 Franks 06/24/2014
If Not X Then GoTo modTermMoveErr_Msg


If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #20052 Franks 07/25/2011
    X% = TERM_PROFIT_SHARING(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
End If

'Ticket #25459 - Terminate ESS and TS employee records as well
'Web Modules Begin
MDIMain.panHelp(0).FloodPercent = 54
X% = TERM_VACTIMEOFF_REQ(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_VACTIMEOFF_REQ_ARCH(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_VACTIMEOFF_REQ_WRK(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_REAUDIT(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_TIMESHEET(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_TIMESHEET_ARCH(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
'Web Modules End

'As Jerry request,We keep Pension employee data in one table, no Active and Term tables
'If glbSQL And glbWFC Then 'Ticket #15537
'    'As Jerry's suggestion, do not put pension tables in Access and Oracle databases
'    x% = TERM_HRP_CREDITED_SERVICE(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'    x% = TERM_HRP_PA_DETAILS(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'    x% = TERM_HRP_PA_MASTER(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'    x% = TERM_HRP_PENSION_BENEFICIARY(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'    x% = TERM_HRP_PENSION_MASTER(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'    x% = TERM_HRP_PENSION_MEMBERSHIP(EEID&)
'    If Not x Then GoTo modTermMoveErr_Msg
'End If

If gsAttachment_DB Then
    X% = TERM_HRDOC_EMP(EEID&)                  'FRANK 01/10/2006
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_JOB_HISTORY(EEID&)          'George 01/19/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_COMMENTS(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HEALTH_SAFETY(EEID&)          'George 02/17/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HEALTH_SAFETY_2(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    
    If glbWSIBModule Then
        X% = TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7(EEID&)
        If Not X Then GoTo modTermMoveErr_Msg
        X% = TERM_HRDOC_OHS_WRITTEN_OFFER(EEID&)
        If Not X Then GoTo modTermMoveErr_Msg
    End If
    
    X% = TERM_HRDOC_COUNSEL(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_PERFORM_HISTORY(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EDSEM(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EDSEM_RETEST(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HREDU(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg '
    X% = TERM_HRDOC_HRDOLENT(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_TRADE(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_ATTENDANCE(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EMP_FLAGS(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    
    'Release 8.1
    X% = TERM_HRDOC_HREMP_OTHER(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
End If '

If glbAxxent Then
    X% = TERM_RSP(EEID&)                  'FRANK 12/22/2000
End If
If glbLinamar Then
    X% = TERM_LN_EMPSKL(EEID&)
End If

MDIMain.panHelp(0).FloodPercent = 55

If Not X Then GoTo modTermMoveErr_Msg

X% = InputHREMPEQU_DOT(EEID&)

gdbAdoIhr001.CommitTrans
'gdbAdoIhr001X.CommitTrans

modTermMove = True

Screen.MousePointer = DEFAULT

Exit Function

modTermMoveErr_Msg:
Screen.MousePointer = DEFAULT

MsgBox TranStr("Problem Creating Audit record - Termination Aborted")

End Function

''Private Sub NukeEE2(EEID As Long)
''Dim snapEETables As New ADODB.Recordset
''Dim SQLQ As String, TabName$
''Dim EEIDAlias$
''
''On Error GoTo NukeEE2_Err
''Dim rsSE As New ADODB.Recordset
''Dim xUserID As String
''rsSE.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE EMPNBR=" & EEID&, gdbAdoIhr001, adOpenStatic
''If Not rsSE.EOF Then
''    xUserID = rsSE("USERID")
''    Call NukeUSERID(xUserID)
''End If
''rsSE.Close
''
''SQLQ = "SELECT * FROM INFO_HR_TABLES "
''SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
''SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
''Ticket #22367, Ticket #20367 - Do not delete employee photo
''SQLQ = SQLQ & " AND Table_Name <>'HR_PHOTO'"
''SQLQ = SQLQ & " AND Table_Name <>'HREEO'"
''
''Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
''Serial 9999 is by default for all standard info:HR table.
''SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
''Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
''SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL)"
''
''snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic
''
''If snapEETables.RecordCount < 1 Then Exit Sub
''snapEETables.MoveFirst
''
''While Not snapEETables.EOF
''    TabName$ = snapEETables("Table_Name")
''    If UCase(Right(TabName$, 3)) <> "WRK" Then
''      EEIDAlias$ = snapEETables("EMPNBR_Alias")
''        If glbVadim And TabName$ = "HRBENFT" Then ' special process for vadim integration
''            gdbAdoIhr001.BeginTrans
''            gdbAdoIhr001.Execute "UPDATE HRBENFT SET BF_LUSER='VADIM_INTEGRATION' WHERE BF_EMPNBR=" & EEID&
''            gdbAdoIhr001.CommitTrans
''        End If
''      Call NukeEERows2(TabName$, EEIDAlias$, EEID&)
''    End If
''    snapEETables.MoveNext
''Wend
''If glbAxxent Then
''    TabName$ = "HRRSP"
''    EEIDAlias$ = "RS_EMPNBR"
''    Call NukeEERows2(TabName$, EEIDAlias$, EEID&)
''End If
''
''snapEETables.Close
''
''Call UpdVacTimeRequest(EEID&, "D")
''
''If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
''    SQLQ = "DELETE FROM HR_OHS_REOCCURENCE WHERE CC_EMPNBR =" & EEID & " "
''    gdbAdoIhr001.Execute SQLQ
''End If
''
''If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
''    SQLQ = "DELETE FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR =" & EEID & " "
''    gdbAdoIhr001.Execute SQLQ
''End If
''
''SQLQ = "DELETE FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR =" & EEID & " "
''gdbAdoIhr001.Execute SQLQ
''
''Exit Sub
''
''NukeEE2_Err:
''glbFrmCaption$ = "Delete Employee"
''glbErrNum& = Err
''
''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
''Call RollBack '29July99 js
''
''End Sub
''
''Private Sub NukeEERows2(TabName As String, EEIDAlias As String, EEID As Long)
'' returns number of records found for ee in table
''Dim Rows%, SQLQ As String
''Dim gdbESS As New ADODB.Connection
''
''On Error GoTo NukeEERows2_Err
''
''If TabName$ = "HREMPEQU" Then
''    Exit Sub
''End If
''
''If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
''    If gdbESS <> "" Then
''        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
''    End If
''End If
''
''SQLQ = "DELETE FROM " & TabName
''SQLQ = SQLQ & " WHERE " & EEIDAlias & " = " & EEID
''
''If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
''    If gdbESS <> "" Then
''        gdbESS.Execute SQLQ
''    End If
''Else
''    gdbAdoIhr001.Execute SQLQ
''End If
''
''Exit Sub
''
''NukeEERows2_Err:
''glbFrmCaption$ = "Nuke Rows"
''glbErrNum& = Err
''
''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", TabName$, "Delete")
''Call RollBack '29July99 js
''
''End Sub

Private Function ReadJob(IJob)
Dim rsTA1 As New ADODB.Recordset

ReadJob = "NO POSITION DESC - " & IJob


rsTA1.Open "HRJOB", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect

'EOF?
rsTA1.MoveFirst
rsTA1.Find "JB_CODE = '" & IJob & "'"

If rsTA1.EOF Then Exit Function

ReadJob = rsTA1("JB_DESCR")

End Function

Private Function READTABLE(Iname, Ikey)
Dim rsTA As New ADODB.Recordset

READTABLE = "No Table Description"

rsTA.Open "HRTABL", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
'rsTA.Index = "HRTABL"
'rsTA.Seek "=", Iname, Ikey
'EOF?
rsTA.MoveFirst
rsTA.Filter = "TB_NAME = '" & Iname & "' and TB_KEY= '" & Ikey & "'"

If rsTA.EOF Then Exit Function

READTABLE = rsTA("TB_DESC")

End Function



Private Sub SELATTWRK()
Dim SQLQ
Dim xFieldList
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "
'Franks - May 7,03 ticket # 4114 - AH_EMPNBR->AD_EMPNBR
SQLQ = SQLQ & " WHERE AD_EMPNBR = " & Val(lblEEID)


gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)

SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
SQLQ = SQLQ & " WHERE AH_EMPNBR = " & Val(lblEEID)

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)


Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub
''Ticket #24184 Franks 11/12/2013use Term_Reviewer_General to replace this
''Private Function Term_Reviewer()
''Dim SQLQDel As String, SQLQCom As String, strTable As String
''Dim dynHRAT As New ADODB.Recordset
''Dim strComm
''
''On Error GoTo Database_Err
''Term_Reviewer = False
''
''strTable = "HR_SUCCESSION"
''SQLQCom = "SELECT EU_REVIEWER FROM HR_SUCCESSION WHERE EU_REVIEWER = " & CLng(lblEEID)
''
''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''
''Screen.MousePointer = HOURGLASS
''
''If dynHRAT.RecordCount >= 1 Then
''    dynHRAT.MoveFirst
''    While Not dynHRAT.EOF
''        dynHRAT("EU_REVIEWER") = 0
''        dynHRAT.Update
''        dynHRAT.MoveNext
''    Wend
''End If
''dynHRAT.Close
''
''Screen.MousePointer = DEFAULT
''
''Term_Reviewer = True
''Exit Function
''
''Database_Err:
''glbFrmCaption$ = Me.Caption
''glbErrNum& = Err
''
''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Reviewer", strTable, "TERMINATE")
''
''End Function

Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "Check this box if some benefits have a different end date than the Termination Date."
    MsgBox MsgStr, vbInformation
End Sub

'''Ticket #24184 Franks 11/12/2013 - use Term_Superv_General to replace this
''Private Function Term_Superv()
'''Laura
''Term_Superv = False
''Dim SQLQDel As String, SQLQCom As String, strTable As String
''Dim dynHRAT As New ADODB.Recordset
''Dim strComm
''
''On Error GoTo Database_Err
'''Set Superv_DB = OpenDatabase(glbIHRDB, False, False)
''
''Screen.MousePointer = HOURGLASS
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''select fields from HR_ATTENDANCE
'''strTable = "HR_ATTENDANCE"
'''SQLQCom = "SELECT * FROM HR_ATTENDANCE WHERE AD_SUPER = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        strComm = dynHRAT("AD_COMM")
'''        If strComm <> "" Then
'''            strComm = strComm & "; "
'''        End If
'''        'dynHRAT.Edit
'''        'dynHRAT("AD_COMM") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
'''        dynHRAT("AD_SUPER") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_ATTENDANCE SET AD_SUPER = 0 WHERE AD_SUPER = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011 - send NULL instead of 0
''SQLQCom = "UPDATE HR_ATTENDANCE SET AD_SUPER = NULL WHERE AD_SUPER = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''select fields from HR_ATTENDANCE_HISTORY
'''strTable = "HR_ATTENDANCE_HISTORY"
'''SQLQCom = "SELECT * FROM HR_ATTENDANCE_HISTORY WHERE AH_SUPER = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        strComm = dynHRAT("AH_COMM")
'''        If strComm <> "" Then
'''            strComm = strComm & "; "
'''        End If
'''        'dynHRAT.Edit
'''        'dynHRAT("AH_COMM") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
'''        dynHRAT("AH_SUPER") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_ATTENDANCE_HISTORY SET AH_SUPER = 0 WHERE AH_SUPER = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_ATTENDANCE_HISTORY SET AH_SUPER = NULL WHERE AH_SUPER = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''select fields from HR_PERFORM_HISTORY
'''strTable = "HR_PERFORM_HISTORY"
'''SQLQCom = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_REPTAU = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        strComm = dynHRAT("PH_COMMENTS")
'''        If strComm <> "" Then
'''            strComm = strComm & "; "
'''        End If
'''        'dynHRAT.Edit
'''        'dynHRAT("PH_COMMENTS") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
'''        dynHRAT("PH_REPTAU") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_PERFORM_HISTORY SET PH_REPTAU = 0 WHERE PH_REPTAU = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_PERFORM_HISTORY SET PH_REPTAU = NULL WHERE PH_REPTAU = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''select fields from HR_JOB_HISTORY
'''strTable = "HR_JOB_HISTORY"
'''SQLQCom = "SELECT * FROM HR_JOB_HISTORY WHERE JH_REPTAU = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        'dynHRAT.Edit
'''        dynHRAT("JH_REPTAU") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = 0 WHERE JH_REPTAU = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = NULL WHERE JH_REPTAU = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU2 = NULL WHERE JH_REPTAU2 = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU3 = NULL WHERE JH_REPTAU3 = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = NULL WHERE JH_REPTAU4 = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''select fields from HR_OCC_HEALTH_SAFETY
'''strTable = "HR_OCC_HEALTH_SAFETY"
'''SQLQCom = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_EMPNOT = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        'dynHRAT.Edit
'''        dynHRAT("EC_EMPNOT") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_EMPNOT = 0 WHERE EC_EMPNOT = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_EMPNOT = NULL WHERE EC_EMPNOT = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
''
'''Hemu - Ticket #16535 - Trying to optimize the process. UPDATE statement is faster than WHILE loop
'''strTable = "HR_OCC_HEALTH_SAFETY"
'''SQLQCom = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_SUPERVISOR = " & CLng(lblEEID)
'''dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''If dynHRAT.RecordCount >= 1 Then
'''    'EOF?
'''    dynHRAT.MoveFirst
'''    While Not dynHRAT.EOF
'''        'dynHRAT.Edit
'''        dynHRAT("EC_SUPERVISOR") = 0
'''        dynHRAT.Update
'''        dynHRAT.MoveNext
'''    Wend
'''End If
'''dynHRAT.Close
'''SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_SUPERVISOR = 0 WHERE EC_SUPERVISOR = " & CLng(lblEEID)
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_SUPERVISOR = NULL WHERE EC_SUPERVISOR = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
'''Ticket #20645 Franks 07/15/2011
''SQLQCom = "UPDATE HR_COUNSEL SET CL_COUBY = NULL WHERE CL_COUBY = " & CLng(lblEEID)
''gdbAdoIhr001.Execute SQLQCom
''
''Screen.MousePointer = DEFAULT
''
''Term_Superv = True
''Exit Function
''
''Database_Err:
''glbFrmCaption$ = Me.Caption
''glbErrNum& = Err
''
''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Superv", strTable, "TERMINATE")
''
''End Function

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEESTATS")
    Call FillMemoFile(SQLQ, "Termination")
End Sub

Private Sub medAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(Me.ActiveControl)
MDIMain.panHelp(2).Caption = " "
End Sub

Private Sub dlpTermDate_LostFocus()
    glbTermDate = dlpTermDate
    If glbWFC Then 'Ticket #19266 Franks 12/13/2010
        If frmLastDay.Visible Then 'Ticket #26308 Franks 11/27/2014
            If Len(xLocLastDay) > 0 Then
                dlpLastDate.Text = xLocLastDay
            Else
                dlpLastDate.Text = dlpTermDate.Text
            End If
        End If
        If dlpDOther2.Visible Then
            If IsDate(dlpTermDate.Text) Then
                If glbTermTran Then 'termination
                    If Len(dlpDOther2.Text) = 0 Then
                        dlpDOther2.Text = dlpTermDate.Text
                    End If
                Else 'transfer out
                    'Ticket #24767 Franks 12/11/2013
                    Call WFCNGSEndDateForTransferOut
                End If
                'Ticket #23247 Franks 07/22/2013
                Call UptData2fromDOT
            End If
        End If
    End If
    
    '8.0 - Ticket #22682 - Update Current Positions with End Date.
    If glbTermTran Then
        If IsDate(dlpTermDate.Text) And Not IsDate(dlpPosEndDate.Text) And chkPosEndDate Then
            dlpPosEndDate.Text = dlpTermDate.Text
        End If
    End If
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

Private Sub EmpWrk()
Dim SQLX, SQLO
Dim xEmpList
Dim xDate1, xDate2
On Error GoTo ERR_EmpWrk
xDate1 = DateAdd("yyyy", -100, Date)
xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

SQLX = "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLX

xEmpList = "(" & lblEEID & ")"
Call glbEmpWrk(xEmpList, xDate1, xDate2)
rr:

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub
ERR_EmpWrk:
If Err = 13 Then
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "EMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub
Private Function TranStr(xStr As String)
TranStr = xStr
If Not glbTermTran Then
    TranStr = Replace(xStr, "Termination", "Transfer")
    TranStr = Replace(TranStr, "Terminate", "Transfer")
    TranStr = Replace(TranStr, "Terminated", "Transferred")
End If
End Function


Private Sub updFollow()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err

rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect

rsTB.AddNew
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = fglbEMPNBR
rsTB("EF_FDATE") = CVDate(dlpTermDate.Text)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(fglbEMPNBR, "ED_ADMINBY", Null)
End If
If glbLinamar Then
    rsTB("EF_FREAS") = "TLAY"
Else
rsTB("EF_FREAS") = "TRAN"
End If
rsTB("EF_COMMENTS") = lblEEName & " was transferred out on " & Format(dlpTermDate.Text, "mmmm dd, yyyy")
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update

rsTB.Close
Msg = "A Follow Up Record was created!"
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub
Private Sub setNewEmpNbrALL()
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xDiv
xDiv = Left(comDIV, 4)
If Not IsNumeric(xDiv) Then Exit Sub
SQLQ = "SELECT EMPNBR FROM "
SQLQ = SQLQ & "(SELECT RIGHT(ED_EMPNBR, LEN(ED_EMPNBR) - 4) AS EMPNBR,ED_EMPNBR AS SORTEMPNBR FROM HREMP WHERE ED_DIV='" & xDiv & "' "
SQLQ = SQLQ & "Union "
SQLQ = SQLQ & "SELECT RIGHT(TL_NEWEMPNBR, LEN(TL_NEWEMPNBR) - 4) AS EMPNBR,TL_NEWEMPNBR AS SORTEMPNBR  FROM LN_TRALOG WHERE TL_NEWDIV='" & xDiv & "') "
SQLQ = SQLQ & "AS UNIONTABLE "
SQLQ = SQLQ & "WHERE CONVERT(INT,EMPNBR)>=1001 "
SQLQ = SQLQ & "ORDER BY SORTEMPNBR "
rsTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

fglbEMPNBR = 1001
Do Until rsTC.EOF
    If fglbEMPNBR <> Val(rsTC("EMPNBR")) Then
        Exit Do
    Else
        fglbEMPNBR = fglbEMPNBR + 1
        rsTC.MoveNext
    End If
Loop
fglbEMPNBR = xDiv & fglbEMPNBR
rsTC.Close
End Sub
Private Sub setNewEmpNbr()
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xDiv
xDiv = Left(comDIV, 3)
If Not IsNumeric(xDiv) Then Exit Sub
SQLQ = "SELECT EMPNBR FROM "
SQLQ = SQLQ & "(SELECT LEFT(ED_EMPNBR, LEN(ED_EMPNBR) - 3) AS EMPNBR,ED_EMPNBR AS SORTEMPNBR FROM HREMP WHERE ED_DIV='" & xDiv & "' AND RIGHT(ED_EMPNBR,3)= '" & xDiv & "' "
SQLQ = SQLQ & "Union "
SQLQ = SQLQ & "SELECT LEFT(TL_NEWEMPNBR, LEN(TL_NEWEMPNBR) - 3) AS EMPNBR,TL_NEWEMPNBR AS SORTEMPNBR  FROM LN_TRALOG WHERE TL_NEWDIV='" & xDiv & "' AND RIGHT(TL_NEWEMPNBR,3)='" & xDiv & "' ) "
SQLQ = SQLQ & "AS UNIONTABLE "
SQLQ = SQLQ & "WHERE CONVERT(INT,EMPNBR)>=1000 "
SQLQ = SQLQ & "ORDER BY SORTEMPNBR "
rsTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

fglbEMPNBR = 1000
Do Until rsTC.EOF
    If fglbEMPNBR <> Val(rsTC("EMPNBR")) Then
        Exit Do
    Else
        fglbEMPNBR = fglbEMPNBR + 1
        rsTC.MoveNext
    End If
Loop
fglbEMPNBR = fglbEMPNBR & xDiv
rsTC.Close
End Sub

''Private Function WFCNonUnion(EmpNbr) As Boolean
''    Dim rsEmp As New ADODB.Recordset, RSTABL As New ADODB.Recordset
''
''    rsEmp.Open "SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
''    If rsEmp.EOF Then
''        WFCNonUnion = False
''        rsEmp.Close
''        Exit Function
''    End If
''    If UCase(rsEmp("ED_ORG")) = "NONE" Or UCase(rsEmp("ED_ORG")) = "EXEC" Then WFCNonUnion = True
''    rsEmp.Close
''End Function

''Private Function IsEmailSetup(EmpNbr) As Boolean
''    Dim rsEmail As New ADODB.Recordset
''
''    rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
''    If rsEmail.EOF Then
''        IsEmailSetup = False
''    Else
''        IsEmailSetup = True
''    End If
''    rsEmail.Close
''End Function
Private Function readFollow()   'Laura on 11/2/97
Dim SQLQ As String
Dim rsTB As New ADODB.Recordset

SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP F"
SQLQ = SQLQ & " INNER JOIN HREMP E ON F.EF_EMPNBR=E.ED_EMPNBR AND F.EF_FDATE=E.ED_UNION"
SQLQ = SQLQ & " WHERE EF_FREAS='TLAY' AND EF_COMPLETED=0 AND EF_EMPNBR=" & glbLEE_ID
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
readFollow = 0
If Not rsTB.EOF Then
    readFollow = rsTB!EF_FOLLOWUP_ID
End If
rsTB.Close

End Function
Private Sub updStatus()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xType
On Error GoTo CrFollow_Err

rsTA.Open "SELECT ED_EMP FROM HREMP WHERE ED_EMP='TEMP' AND ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Sub

SQLQ = "SELECT * FROM HRSTATUS "
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
rsTB.AddNew
rsTB("SC_COMPNO") = "001"
rsTB("SC_EMPNBR") = glbLEE_ID
rsTB("SC_EMP_TABL") = "EDEM"
rsTB("SC_REASON_TABL") = "SCRE"

If IsDate(dlpTermDate.Text) Then rsTB("SC_FDATE") = dlpTermDate.Text
rsTB("SC_OLDEMP") = rsTA!ED_EMP
rsTB("SC_NEWEMP") = rsTA!ED_EMP
rsTB("SC_REASON") = "TERM"

rsTB("SC_FOLLOWID") = readFollow
rsTB("SC_JOB") = ReadJobCode
rsTB("SC_LDATE") = Date
rsTB("SC_LTIME") = Time$
rsTB("SC_LUSER") = glbUserID
rsTB.Update
rsTB.Close
rsTA.Close
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "STATUS", "HRSTATUS", "UPDATE TABLE")
Resume Next

End Sub
Private Function ReadJobCode()
Dim rsTA As New ADODB.Recordset
Dim IJob
ReadJobCode = ""

rsTA.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Function
ReadJobCode = rsTA("JH_JOB")
rsTA.Close

End Function


Private Sub SaveDelayInfo()
Dim rsTA As New ADODB.Recordset, SQLQ
rsTA.Open "SELECT * FROM HR_WILL_TERM WHERE TL_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockPessimistic
If rsTA.EOF Then rsTA.AddNew
rsTA!TL_COMPNO = "001"
rsTA!TL_EMPNBR = glbLEE_ID
rsTA!TL_TYPE = IIf(glbTermTran, "T", "R")
rsTA!TL_DOT = dlpTermDate
rsTA!TL_Reason_Tabl = "TERM"
rsTA!TL_Reason = clpCode(1)
rsTA!TL_Rehire = IIf(chkRehire, 1, 0)
rsTA!TL_SUMATT = IIf(chkSum, 1, 0)
rsTA!TL_COMMENTS = txtComments
If glbTermTran Then
    rsTA!TL_TYPE = "T"
Else
    rsTA!TL_TYPE = "R"
    rsTA!TL_TODIV = Left(comDIV, 3)
    rsTA!TL_TOEMPNBR = getEmpnbr(lblEMPNo)
End If
rsTA!TL_LDATE = Date
rsTA!TL_LTIME = Time$
rsTA!TL_LUSER = glbUserID
rsTA.Update
If glbTermTran Then
    MsgBox "The employee will be automatically terminated on " & Format(dlpTermDate, "mmm dd, yyyy")
Else
    MsgBox "The employee will be automatically transferred out on " & Format(dlpTermDate, "mmm dd, yyyy")
End If
glbLEE_ID = 0
dlpTermDate.Text = ""
clpCode(1).Text = ""

'Ticket #21238 - County of Oxford
If glbCompSerial = "S/N - 2259W" Then
    chkRehire.Value = False
    lblRehire.Caption = "No"
Else
    chkRehire.Value = True
End If

txtComments.Text = ""
Unload frmEEBASIC 'Add by Frank Aug 17, 01
Call UnloadFrms
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
UpdateRight = gSec_Upd_Terminations
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
Printable = False
End Property

Public Sub SET_UP_MODE()
    Dim TF As Boolean
    Dim UpdateState As UpdateStateEnum
    
    TF = True
    UpdateState = OPENING
    
    Call set_Buttons(UpdateState)
    
    If Not UpdateRight Then TF = False
End Sub

Sub cmdCancel_Click()
    Dim X As Integer
    
    fglbNew = False
    
    dlpTermDate.Text = ""
    clpCode(1).Text = ""
    txtComments.Text = ""
    
    For X = 0 To 5
        chkTermRpts(X).Value = True
    Next
    
    chkSum.Value = False
    
    'Ticket #21238 - County of Oxford
    If glbCompSerial = "S/N - 2259W" Then
        chkRehire.Value = False
        lblRehire.Caption = "No"
    Else
        chkRehire.Value = True
    End If
        
    Call SET_UP_MODE
    
    Exit Sub

End Sub

Private Sub UpdPositionCCAC()
Dim rsOC As New ADODB.Recordset
Dim SQLQ
If glbOttawaCCAC Then
    SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_EMPNBR =" & glbLEE_ID
    rsOC.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsOC.EOF Then
        rsOC("PC_EMPNBR") = Null
        rsOC.Update
    End If
    rsOC.Close
    
End If
End Sub

Private Sub UpdEmpType(xValue)
Dim SQLQ
    SQLQ = "UPDATE HREMP SET ED_EMPTYPE = '" & xValue & "' WHERE ED_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub UpdEmpStatus(xValue)
Dim SQLQ
    SQLQ = "UPDATE HREMP SET ED_EMP = '" & xValue & "' WHERE ED_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub ClearEmpBenefitGroup()
Dim SQLQ
    SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = '' WHERE ED_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub UpdPositionMulti()
Dim rsJH As New ADODB.Recordset
Dim rsSH As New ADODB.Recordset
Dim rsPH As New ADODB.Recordset
Dim SQLQ
If glbMulti Then
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR =" & glbLEE_ID
    rsJH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsJH.EOF
        rsJH("JH_CURRENT") = 0
        rsJH("JH_ENDDATE") = dlpTermDate
        rsJH("JH_ENDREAS") = clpCode(1)
        rsJH.Update
        rsJH.MoveNext
    Loop
    rsJH.Close
    
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR =" & glbLEE_ID
    rsSH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsSH.EOF
        rsSH("SH_CURRENT") = 0
        rsSH.Update
        rsSH.MoveNext
    Loop
    rsSH.Close
    
    SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_CURRENT<>0 AND PH_EMPNBR =" & glbLEE_ID
    rsPH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsPH.EOF
        rsPH("PH_CURRENT") = 0
        rsPH.Update
        rsPH.MoveNext
    Loop
    rsPH.Close
    
End If
End Sub

Private Sub UpdPositionEndDate()
    Dim rsJH As New ADODB.Recordset
    Dim SQLQ
    
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR =" & glbLEE_ID
    rsJH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsJH.EOF
        'rsJH("JH_CURRENT") = 0 'Will cause an issue when running terminated reports as the record will not show up.
        rsJH("JH_ENDDATE") = dlpTermDate
        'rsJH("JH_ENDREAS") = clpCode(1)
        rsJH.Update
        rsJH.MoveNext
    Loop
    rsJH.Close
    Set rsJH = Nothing
End Sub

Private Sub UpdPaymentTypeVadim()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
If glbVadim Then
    If Vadim_PayType_field = "" Then Exit Sub
    SQLQ = "SELECT " & Vadim_PayType_field & " FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsEmp.EOF Then

    If glbChgTermReason = "LO" Or glbChgTermReason = "LAYO" Then
        rsEmp(Vadim_PayType_field) = "L"
        
        'City of Kawartha Lakes - When Terminate - they pass T code. If employee's benefits
        'are continued they do not terminate employee but instead just changes the Payment Type
        'R or L.
        If glbCompSerial = "S/N - 2363W" Then
            rsEmp(Vadim_PayType_field) = "T"
        End If
        
    ElseIf glbChgTermReason = "RETI" Then
        rsEmp(Vadim_PayType_field) = "R"
    
        'City of Kawartha Lakes - When Terminate - they pass T code. If employee's benefits
        'are continued they do not terminate employee but instead just changes the Payment Type
        'R or L.
        If glbCompSerial = "S/N - 2363W" Then
            rsEmp(Vadim_PayType_field) = "T"
        End If
    Else
        rsEmp(Vadim_PayType_field) = "T"
    End If
    rsEmp.Update
    End If
    rsEmp.Close
End If
End Sub

Sub SubPicture()
Dim xPIC
Dim Msg As String
Dim xHeight, xWidth

On Error GoTo cmdPic_ERR

If glbtermopen Then Exit Sub

'8.0 - Ticket #22682 - Photo from the folder now
'If glbSQL Or glbOracle Then
'    If cmdPhoto.Caption = "&Photo Off" Then
'        picPhoto.Visible = False
'        PicNotF.Visible = False
'        cmdPhoto.Caption = "&Photo"
'    Else
'        picPhoto.Visible = False
'        PicNotF.Visible = True
'        cmdPhoto.Caption = "&Photo Off"
'        Call FillPhoto(Val(lblEEID))
'    End If
'Else
    If Len(glbPicDir) < 1 Then
        picPhoto.Visible = False
        Exit Sub
    End If
    If cmdPhoto.Caption = "&Photo Off" Then
        picPhoto.Visible = False
        PicNotF.Visible = False
        picPhoto = LoadPicture()
        cmdPhoto.Caption = "&Photo"
    Else
        picPhoto.Visible = False
        PicNotF.Visible = True
        cmdPhoto.Caption = "&Photo Off"
        Call LoadPhoto(Val(lblEEID))
    End If
'End If

Exit Sub

cmdPic_ERR:
If Err Then
  PicNotF.Visible = True
  Exit Sub
End If

End Sub

Private Function FillPhoto(zEMPNBR As Long)
    On Error GoTo ErrHandler:
    Dim rsPHOTO As New ADODB.Recordset
    Dim byteChunk() As Byte

    Dim Offset As Long
    Dim Totalsize As Long
    Dim Remainder As Long

    Dim FieldSize As Long
    Dim FileNumber As Integer
    Const HeaderSize As Long = 78
    Const ChunkSize As Long = 100
    Dim TempFile As String
    Dim TempDir As String * 255
    
    GetTempPath 255, TempDir
    TempFile = Replace(Replace(TempDir, Chr(0), "") & "\tempfile.tmp", "\\", "\")
    
    picPhoto.Picture = Nothing
    If zEMPNBR = 0 Then Exit Function
    rsPHOTO.Open "select * from HR_PHOTO WHERE PT_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsPHOTO.EOF Then Exit Function
    
    
    FileNumber = FreeFile
    Open TempFile For Binary Access Write As FileNumber
    
    ReDim byteChunk(rsPHOTO("PT_PHOTO").ActualSize)
    byteChunk() = rsPHOTO("PT_PHOTO").GetChunk(rsPHOTO("PT_PHOTO").ActualSize)
    Put FileNumber, , byteChunk()

    Close FileNumber
    picPhoto.Picture = LoadPicture(TempFile)
    Kill (TempFile)
    rsPHOTO.Close
    Dim xHeight, xWidth
    picPhoto.Stretch = False
    xHeight = picPhoto.Height
    xWidth = picPhoto.Width
    picPhoto.Stretch = True
    picPhoto.Height = 2325
    picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
    picPhoto.Stretch = True
    picPhoto.Visible = True
    PicNotF.Visible = False
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, , "Error "
    
End Function

Private Function LoadPhoto(zEMPNBR As Long)
Dim xHeight, xWidth

glbPicBMP = glbPicDir & zEMPNBR & ".JPG"

'Hemu
If Not IsNull(glbPicBMP) Then
    If Not (Dir(glbPicBMP) = "") Then
        picPhoto = LoadPicture(glbPicBMP)
    Else
        Exit Function
    End If
Else
    Exit Function
End If
'If Not IsNull(glbPicBMP) Then picPhoto = LoadPicture(glbPicBMP)
'Hemu

picPhoto.Stretch = False
xHeight = picPhoto.Height
xWidth = picPhoto.Width
picPhoto.Stretch = True
picPhoto.Height = 2325
picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
picPhoto.Stretch = True
picPhoto.Visible = True
PicNotF.Visible = False
End Function

Private Sub PhotoFormLoad()
Dim xPIC

    '8.0 - Ticket #22682 - Move Photo out of database into a folder ---------------------------------------------------
    'Get Photo Path
    If gsEMPLOYEEPHOTO Then
        xPIC = GetComPreferEmail("EMPLOYEEPHOTOPATH")
        If Len(xPIC) > 0 And Right(xPIC, 1) <> "\" Then xPIC = xPIC & "\"
    End If
    'xPIC = glbIHRREPORTS & "IHRPICS.MTR"   'Ticket #22682
    
    If xPIC = "" Then   'Ticket #22682
    'If (Dir(xPIC) = "" And Not glbOracle And Not glbSQL) Or glbtermopen Then
        PicNotF.Visible = False
        cmdPhoto.Enabled = False 'Jaddy 10/28/99
        picPhoto.Visible = False
        glbPicDir = ""
        cmdPhoto.Caption = "&Photo"
    Else
        PicNotF.Visible = True
        cmdPhoto.Enabled = True 'Jaddy 10/28/99
        picPhoto.Visible = False
        'glbPicDir = glbIHRREPORTS  'Ticket #22682
        glbPicDir = xPIC
    End If

    'Ticket #22682
'    If glbSQL Or glbOracle Then
'        If cmdPhoto.Caption = "&Photo Off" Then
'            picPhoto.Visible = False
'            PicNotF.Visible = True
'            Call FillPhoto(Val(glbLEE_ID))
'        Else
'            picPhoto.Visible = False
'            PicNotF.Visible = False
'        End If
'    Else
        If Len(glbPicDir) < 1 Then
            picPhoto.Visible = False
        Else
            If cmdPhoto.Caption = "&Photo Off" Then
                picPhoto.Visible = False
                PicNotF.Visible = True
                Call LoadPhoto(Val(glbLEE_ID))
            Else
                picPhoto.Visible = False
                PicNotF.Visible = False
            End If
        End If
'    End If
    
End Sub
''George on Jan 10,2005
''Ticket #24184 Franks 11/12/2013
'Private Sub UpdEHScorrective()
'    Dim SQLQ
'    Dim xName
'    If InStr(1, lblEEName.Caption, "'") > 0 Then
'        xName = Replace(lblEEName.Caption, "'", "''")
'        If glbLinamar Then
'            SQLQ = "update HR_OHS_CORRECTIVE set CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & glbLEE_ID
'        Else
'            SQLQ = "update HR_OHS_CORRECTIVE set CR_ASSIGNED  = Null,CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & glbLEE_ID
'        End If
'    Else
'        If glbLinamar Then
'            SQLQ = "update HR_OHS_CORRECTIVE set CR_TERM_EMPNAME ='" & lblEEName.Caption & "' WHERE CR_ASSIGNED =" & glbLEE_ID
'        Else
'            SQLQ = "update HR_OHS_CORRECTIVE set CR_ASSIGNED  = Null,CR_TERM_EMPNAME ='" & lblEEName.Caption & "' WHERE CR_ASSIGNED =" & glbLEE_ID
'        End If
'    End If
'    gdbAdoIhr001.Execute SQLQ
'End Sub

Public Sub imgEmail_Click()
    If Not glbWFC And Not (glbCompSerial = "S/N - 2382W") Then '2382 - Samuel
        Call cmdEmail_Click
    End If
End Sub

Private Sub LoadStatusFlag()

cboStatFlag3.AddItem "1 - Fired for Cause"
cboStatFlag3.AddItem "2 - Laid Off - do no use a term date"
cboStatFlag3.AddItem "4 - Retirement"
cboStatFlag3.AddItem "5 - Permanent Disability"
cboStatFlag3.AddItem "6 - Death"
cboStatFlag3.AddItem "9 - Transferred to another payroll code"
cboStatFlag3.AddItem "L - Leave of Absence paid or unpaid"

End Sub
'''Ticket #24184 Franks 11/12/2013 replace this function with AUDIT_MANULIFE_TRANS_TermTransferOut
''Private Function AUDIT_MANULIFE_TRANS() 'No AU_CEASEDATE in HRAUDIT, Jerry said we will add it in next release
''Dim rsTA As New ADODB.Recordset
''Dim rsTB As New ADODB.Recordset
''Dim rsBene As New ADODB.Recordset
''Dim rsDepend As New ADODB.Recordset
''Dim xADD As Boolean, xPT As String, xDiv As String
''Dim strFields As String
''Dim SQLQ As String
'''''On Error GoTo AUDIT_ERR
''AUDIT_MANULIFE_TRANS = False
''
'''BENEFIT End Date
''If Len(locCertNo) = 0 Or Len(glbChgBenTermDate) = 0 Then
''    Exit Function
''End If
''
''rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
''If rsTB.EOF Then
''    rsTB.Close:    GoTo MODNOUPD_Den
''End If
''If IsNull(rsTB("ED_USER_TEXT1")) Then 'Certificate #
''    rsTB.Close:    GoTo MODNOUPD_Den
''Else
''    If Len(Trim(rsTB("ED_USER_TEXT1"))) = 0 Then
''        rsTB.Close:    GoTo MODNOUPD_Den
''    End If
''End If
''
'''Benefits
''SQLQ = "SELECT * FROM HRBENFT WHERE NOT(BF_POLICY IS NULL) AND BF_EMPNBR = " & glbLEE_ID
''rsBene.Open SQLQ, gdbAdoIhr001, adOpenStatic
''If rsBene.EOF Then
''    rsBene.Close
''    GoTo MODNOUPD_Ben 'Exit Function
''End If
''
''
''Do While Not rsBene.EOF
''    If Len(rsBene("BF_POLICY")) > 0 Then
''        If Not IsDate(rsBene("BF_CEASEDATE")) Then 'No Benefit End Date
''            If rsTA.State <> 0 Then rsTA.Close
''            rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
''
''            rsTA.AddNew
''            rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
''            rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
''            rsTA("MT_PT_TABL") = "EDPT"
''            rsTA("MT_TYPE") = "T"
''            rsTA("MT_BENEFIT") = rsBene("BF_BCODE")
''            rsTA("MT_EDATE") = rsBene("BF_EDATE")
''            rsTA("MT_CEASEDATE") = glbChgBenTermDate
''            rsTA("MT_COVER") = rsBene("BF_COVER")
''            rsTA("MT_COMPNO") = "001"
''            rsTA("MT_EMPNBR") = glbLEE_ID
''            rsTA("MT_POLICY_NO") = rsBene("BF_POLICY")
''            rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
''            rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
''            rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
''            rsTA("MT_UPLOAD") = "N"
''            rsTA("MT_LUSER") = glbUserID
''            If CVDate(glbChgBenTermDate) < CVDate(Date) Then 'Ticket #14867
''                rsTA("MT_LDATE") = Date
''            Else
''                rsTA("MT_LDATE") = Format(glbChgBenTermDate, "SHORT DATE")
''            End If
''            rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
''            rsTA("MT_LTIME") = Time$
''
''            rsTA.Update
''        End If
''    End If
''    rsBene.MoveNext
''Loop
''rsBene.Close
''
''MODNOUPD_Ben:
''
''SQLQ = "SELECT * FROM HRDEPEND WHERE DP_EMPNBR = " & glbLEE_ID
''rsDepend.Open SQLQ, gdbAdoIhr001, adOpenStatic
''If rsDepend.EOF Then
''    rsDepend.Close
''    GoTo MODNOUPD_Den 'Exit Function
''End If
''
''Do While Not rsDepend.EOF
''    If Not IsDate(rsDepend("DP_EDATE")) Then 'No Benefit End Date
''        If rsTA.State <> 0 Then rsTA.Close
''        rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
''
''        rsTA.AddNew
''        rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
''        rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
''        rsTA("MT_PT_TABL") = "EDPT"
''        rsTA("MT_TYPE") = "T"
''        rsTA("MT_DEPFNAME") = rsDepend("Dp_FName")
''        rsTA("MT_DEPSNAME") = rsDepend("DP_SNAME")
''        rsTA("MT_DEPSEX") = rsDepend("DP_SEX")
''        rsTA("MT_DEPDOB") = rsDepend("DP_DOB")
''        rsTA("MT_DEPRELATE") = rsDepend("DP_RELATE")
''        rsTA("MT_DEPSMOKER") = rsDepend("DP_SMOKER")
''        rsTA("MT_DEPSTATUS") = rsDepend("DP_STATUS")
''        rsTA("MT_DEPSIN") = rsDepend("DP_SIN")
''        rsTA("MT_DEPSDATE") = rsDepend("DP_SDATE")
''        rsTA("MT_DEPEDATE") = glbChgBenTermDate
''        rsTA("MT_DENTAL") = rsDepend("DP_DENTAL")
''        rsTA("MT_MEDICAL") = rsDepend("DP_MEDICAL")
''        rsTA("MT_OTHER") = rsDepend("DP_OTHER")
''        rsTA("MT_COMPNO") = "001"
''        rsTA("MT_EMPNBR") = glbLEE_ID
''        rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
''        rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
''        rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
''        rsTA("MT_UPLOAD") = "N"
''        rsTA("MT_LUSER") = glbUserID
''        If CVDate(glbChgBenTermDate) < CVDate(Date) Then 'Ticket #14867
''            rsTA("MT_LDATE") = Date
''        Else
''            rsTA("MT_LDATE") = Format(glbChgBenTermDate, "SHORT DATE")
''        End If
''        rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
''        rsTA("MT_LTIME") = Time$
''        rsTA.Update
''    End If
''    rsDepend.MoveNext
''Loop
''rsDepend.Close
''
''MODNOUPD_Den:
''
''AUDIT_MANULIFE_TRANS = True
''Exit Function
''AUDIT_ERR:
''
''glbFrmCaption$ = Me.Caption
''glbErrNum& = Err
''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT RECORD", "MANULIFE AUDIT FILE", "UPDATE")
''If gintRollBack% = False Then Resume Next Else Unload Me
''
''End Function

Private Function AUDIT_BenefitEndDate()
    Dim rsAudit As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsBene As New ADODB.Recordset
    Dim SQLQ As String
    Dim strFields As String
    Dim xPT As String, xDiv As String, xPayrollID As String

    On Error GoTo AUDIT_BenefitEndDate_ERR
    
    AUDIT_BenefitEndDate = False
    
    rsHREmp.Open "SELECT ED_PT,ED_DIV,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsHREmp.EOF Then
        If IsNull(rsHREmp("ED_PT")) Then
            xPT = ""
        Else
            xPT = rsHREmp("ED_PT")
        End If
        If IsNull(rsHREmp("ED_DIV")) Then
            xDiv = ""
        Else
            xDiv = rsHREmp("ED_DIV")
        End If
        If IsNull(rsHREmp("ED_PAYROLL_ID")) Then
            xPayrollID = ""
        Else
            xPayrollID = rsHREmp("ED_PAYROLL_ID")
        End If
    Else
        xPT = ""
        xDiv = ""
        xPayrollID = ""
    End If
    
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
    strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE," ', AU_MAXDOL, AU_PPAMT, "
    strFields = strFields & "AU_BCODE,AU_POLICY,AU_CEASEDATE,"  'AU_MTHCCOST, AU_MTHECOST, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC,
    'strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST"
    strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE" ',AU_OLDLOC,AU_OLDWHRS "
    
    SQLQ = "SELECT * FROM HRBENFT WHERE (BF_CEASEDATE IS NULL) AND BF_EMPNBR = " & glbLEE_ID
    rsBene.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBene.EOF Then
        Do While Not rsBene.EOF
            If Not IsDate(rsBene("BF_CEASEDATE")) Then 'No Benefit End Date
                If rsAudit.State <> 0 Then rsAudit.Close
            
                rsAudit.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
                rsAudit.AddNew
                rsAudit("AU_LOC_TABL") = "EDLC": rsAudit("AU_SECTION_TABL") = "EDSE": rsAudit("AU_EMP_TABL") = "EDEM": rsAudit("AU_SUPCODE_TABL") = "EDSP"
                rsAudit("AU_ORG_TABL") = "EDOR": rsAudit("AU_PAYP_TABL") = "SDPP": rsAudit("AU_BCODE_TABL") = "BNCD": rsAudit("AU_TREAS_TABL") = "TERM"
                rsAudit("AU_DOLENT_TABL") = "EDOL": rsAudit("AU_EARN_TABL") = "EARN"
                rsAudit("AU_NEWEMP") = "N"
                rsAudit("AU_PTUPL") = xPT
                rsAudit("AU_DIVUPL") = xDiv
                rsAudit("AU_COMPNO") = "001"
                rsAudit("AU_EMPNBR") = glbLEE_ID
                rsAudit("AU_PAYROLL_ID") = xPayrollID
                rsAudit("AU_BCODE") = rsBene("BF_BCODE")
                rsAudit("AU_EDATE") = rsBene("BF_EDATE")
                rsAudit("AU_COVER") = rsBene("BF_COVER")
                rsAudit("AU_POLICY") = rsBene("BF_POLICY")
                rsAudit("AU_CEASEDATE") = dlpBenCeaseDate
                
                If CVDate(dlpBenCeaseDate) < CVDate(Date) Then 'Ticket #14867
                    rsAudit("AU_LDATE") = Date
                Else
                    rsAudit("AU_LDATE") = Format(dlpBenCeaseDate, "SHORT DATE")
                End If
                rsAudit("AU_LUSER") = glbUserID
                rsAudit("AU_LTIME") = Time$
                rsAudit("AU_UPLOAD") = "N"
                rsAudit("AU_TYPE") = "T"
                rsAudit.Update
                
                rsBene("BF_CEASEDATE") = dlpBenCeaseDate
                rsBene.Update
            End If
            rsBene.MoveNext
        Loop
    End If
    rsBene.Close
    Set rsBene = Nothing

Audit_BenUpd:
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    AUDIT_BenefitEndDate = True

Exit Function

AUDIT_BenefitEndDate_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING BENEFIT END DATE AUDIT RECORD", "AUDIT FILE", "ADD")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function


Private Sub AUDIT_GWL_TRANS()
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpID
Dim xForm As String
Dim xTranType
Dim xChgType
Dim xEDate, xDate1, xDate2
Dim xLDate
Dim xBenGroup

On Error GoTo AUDIT_ERR

    If Not glbIsGWL Then Exit Sub
    SQLQ = "SELECT ED_EMPNBR, ED_BENEFIT_GROUP, ED_DOH, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_BENEFIT_GROUP")) Then xBenGroup = "" Else xBenGroup = rsEmpee("ED_BENEFIT_GROUP")
    End If
    'rsEmpee.Close
    'No Benefit Group Code, skip
    If Len(xBenGroup) = 0 Then Exit Sub
    xEmpID = glbLEE_ID
    xTranType = "T"
    If clpCode(1).Text = "RETI" Then
        xChgType = "Retirement"
    Else
        xChgType = "Termination"
    End If
    If IsDate(dlpBenCeaseDate.Text) Then
        xEDate = dlpBenCeaseDate.Text
    Else
        xEDate = dlpTermDate.Text
    End If
    xForm = "Termination"
    xLDate = Date
    If CVDate(xEDate) > CVDate(xLDate) Then
        xLDate = xEDate
    End If
    rsEmpee.Close
    
    'GWL field changes --------------------------------------
    Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Term Reason", "", clpCode(1).Text, xLDate)

    Exit Sub

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING GWL AUDIT RECORD", "GWL AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub
'''Ticket #24184 Franks 11/12/2013 - use "WFC_NGS_Trans_TermTransferOut" replace this
''Private Sub WFC_NGS_Trans(xType) '#19266
''Dim rsEmpee As New ADODB.Recordset
''Dim rsEmpOther As New ADODB.Recordset
''Dim SQLQ As String
''Dim xUnion As String
''Dim xSalHly As String
''Dim xInSubGrp As String
''Dim xLDate
''Dim xNGSStart
''Dim xCurPlant, xToPlant 'Ticket #23501 Franks 04/02/2013
''
''    If Not glbNGS_OnFlag Then
''        Exit Sub
''    End If
''
''    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
''    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
''    If rsEmpee.EOF Then
''        Exit Sub
''    Else
''        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
''        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
''        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
''        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
''    End If
''    rsEmpee.Close
''
''    'No NGS Sub Group, skip
''    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub
''
''    'Ticket #20385 Franks 05/31/2011
''    'xLDate = dlpTermDate.Text 'Date
''    xLDate = Date
''
''    xNGSStart = ""
''    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
''    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
''    If Not rsEmpOther.EOF Then
''        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
''            xNGSStart = rsEmpOther("ER_OTHERDATE1")
''        End If
''    End If
''    rsEmpOther.Close
''
''    ''Ticket #20385 Franks 05/31/2011
''    ''No NGS Effective Date, skip
''    'If Len(xNGSStart) = 0 Then Exit Sub
''
''    If glbUNION = "NONE" Or glbUNION = "EXEC" Then
''        xSalHly = "Y"
''    Else
''        xSalHly = "N"
''    End If
''
''    If xType = "Transfer Out" Then
''        If Len(clpCode(3).Text) > 0 Then 'Ticket #21677 Franks 03/14/2012
''            xInSubGrp = getNGSSubGrpFromMatrix(Trim(Left(comDIV.Text, 4)), clpCode(3).Text)
''        Else
''            xInSubGrp = getNGSSubGrpFromMatrix(Trim(Left(comDIV.Text, 4)), glbUNION)
''        End If
''        If Len(xInSubGrp) = 0 Then
''            'Ticket #23501 Franks 04/02/2013
''            '"   If the Plant in the Transfer To Division equals the Plant from the employee's record
''            'do not create the NGS Audit record and do not populate the NGS End Date field.
''            xCurPlant = getSectionByDiv(glbEmpDiv)
''            xToPlant = getSectionByDiv(Left(comDIV.Text, 4))
''            If xCurPlant = xCurPlant Then
''                'transfer between same plant, do not change NGS
''            Else
''                Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", CVDate(dlpTermDate.Text))
''                'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(dlpTermDate.Text), xLDate)
''                'Ticket #22409 Franks 08/16/2012 send NGS End Date only when the employee transfer out NGS group
''                Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(dlpTermDate.Text), xLDate)
''            End If
''        End If
''        'Ticket #22409 Franks 08/16/2012 do not send NGS End Date between unions
''        ''Ticket #21822 Franks 04/10/2012 - Send NGS End Date to NGS for transfer between unions or Div
''        'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(dlpTermDate.Text), xLDate)
''    End If
''    If xType = "Termination" Then
''        If IsDate(dlpDOther2.Text) Then
''            Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", CVDate(dlpDOther2.Text))
''            'Call NGSAuditAdd(glbLEE_ID, "M", "Termination", "Transfer Out Date", "", CVDate(dlpTermDate.Text), xLDate)
''            'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", "To Division", glbEmpDiv, Trim(Left(comDIV.Text, 4)), xLDate)
''            Call NGSAuditAdd(glbLEE_ID, "M", "Termination", lStr("Other Date 2"), "", CVDate(dlpDOther2.Text), xLDate)
''        End If
''    End If
''End Sub

Private Sub WFCOther2Screen(xEmpNo)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart

    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    'Ticket #19678 Franks 01/21/2011
    lblTitle(5).Enabled = True
    dlpBenCeaseDate.Enabled = True
    
    lbOtherDate2.Visible = False
    dlpDOther2.Visible = False
    
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2,ED_LDAY FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
        If IsNull(rsEmpee("ED_LDAY")) Then xLocLastDay = "" Else xLocLastDay = rsEmpee("ED_LDAY")
    End If
    rsEmpee.Close
    If glbTermTran Then 'Ticket #26308 Franks 11/27/2014
        dlpLastDate.Text = xLocLastDay
    End If
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    
    'Ticket #20385 Franks 05/31/2011
    ''No NGS Effective Date, skip
    'If Len(xNGSStart) = 0 Then Exit Sub
    
    lbOtherDate2.Caption = lStr("Other Date 2")
    If glbTermTran Then
        If medAmount.Visible Then
            lbOtherDate2.Top = 1270 + 310
            dlpDOther2.Top = 1270 + 310
        Else
            lbOtherDate2.Top = 1270
            dlpDOther2.Top = 1270
        End If
        lbOtherDate2.Left = lblTitle(0).Left
        dlpDOther2.Left = dlpTermDate.Left
        lbOtherDate2.Visible = True
        dlpDOther2.Visible = True
    Else 'transfer Out
        'Ticket #24767 Franks 12/11/2013
        lbOtherDate2.Top = 1930
        dlpDOther2.Top = 1930
        lbOtherDate2.FontBold = False
        lbOtherDate2.Left = lblTitle(0).Left
        dlpDOther2.Left = dlpTermDate.Left
        lbOtherDate2.Visible = True
        dlpDOther2.Visible = True
        
        lblTitle(5).Visible = False
        dlpBenCeaseDate.Visible = False
    End If
    
    'Ticket #19678 Franks 01/21/2011
    'If NGS employee, disable End Date on the Termination screen. - Jerry
    If IsDate(xNGSStart) Then
        lblTitle(5).Enabled = False
        dlpBenCeaseDate.Enabled = False
        Call WFCBenListScreen(glbLEE_ID) 'Ticket #23247 Franks 07/22/2013
    End If
End Sub

Private Sub EEO_Process() 'Ticket #20270 Franks 05/05/2011
    If glbEmpCountry = "U.S.A." Then
        Call uptEEO_Fields(glbLEE_ID, "Delete")
    End If
End Sub

Private Sub CheckReptAuth() 'Ticket #20885 Franks 11/18/2011 for Samuel
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
Dim SQLQ As String

    xFlag1 = False
    'check if this employee is a Reporting Authority
    If IsReportAuth(glbLEE_ID) Then
        xFlag1 = True
    End If

    If xFlag1 Then
        xMsg = "This employee has been assigned as a Reporting Authority on other employee files." & Chr(10)
        xMsg = xMsg & "This termination affects the Reporting Authority structures" & Chr(10)
        xMsg = xMsg & "For all employees affected, this employee will be automatically removed from their Reporting Authorities."
        'frmMsgYesNoUn.lblMsg.Caption = xMsg
        'frmMsgYesNoUn.lblMsg.Alignment = 0
        'frmMsgYesNoUn.Show 1
        MsgBox xMsg
        'If glbMsgCustomVal = 1 Or glbMsgCustomVal = 3 Then
            'create a report to show the employee list
            Call CreateEmpList4ReportAuth(glbLEE_ID)
            'show the report - begin
            Me.vbxCrystal.Reset
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
            End If
            'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & lblEEName & "'"
            'Ticket #21669 Franks 03/01/2012
            xMsg = Replace(lblEEName, "'", "''")
            Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
            Me.vbxCrystal.Connect = RptODBC_SQL
            Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & lblEEName
            Me.vbxCrystal.Destination = 0
            Me.vbxCrystal.Action = 1
            Me.vbxCrystal.Reset
            'show the report - end
        'End If
        
        'this employee will be automatically removed from their Reporting Authorities.
        'Reporting Authority 1
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU = " & glbLEE_ID & " "
        gdbAdoIhr001.Execute SQLQ
        'Reporting Authority 2
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU2 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU2 = " & glbLEE_ID & " "
        gdbAdoIhr001.Execute SQLQ
        'Reporting Authority 3
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU3 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU3 = " & glbLEE_ID & " "
        gdbAdoIhr001.Execute SQLQ
        'Reporting Authority 4
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU4 = " & glbLEE_ID & " "
        gdbAdoIhr001.Execute SQLQ

    End If
    
End Sub

'''Ticket #24184 Franks 11/12/2013 - replace it with locWFCUpdPAMaster_TermTransferOut
''Private Sub locWFCUpdPAMaster(xEmpNo, xTranDate, xSIN) ', xNewDiv)
'''o   If the Union Code changes to "NONE" or "EXEC", a PA Master must be created
'''"   Earned Pension is calculated using the Hourly Year End Pension & PA Update rules. (Credited Service months * Benefit Rate)
'''Frank Note: the Union Transfer Out create the Pension Master with status code X first
'''so this PA Master update function will reuse the same Earning Pension from Pension Master
''Dim rsPen As New ADODB.Recordset
''Dim rsPAMaster As New ADODB.Recordset
''Dim SQLQ As String
''Dim xYear
''Dim xEarnPen, xTotal
''    If Len(xSIN) = 0 Then Exit Sub
''    If Not IsDate(xTranDate) Then Exit Sub
''
''    xYear = Year(CVDate(xTranDate))
''    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
''    SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
''    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
''    SQLQ = SQLQ & "AND PE_DB_STATUS = 'X' "
''    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
''    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC"
''    rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
''    If Not rsPen.EOF Then
''        xEarnPen = 0
''        If Not IsNull(rsPen("PE_CREDITED_SERV")) And Not IsNull(rsPen("PE_BENEFIT_RATE")) Then
''            xEarnPen = rsPen("PE_CREDITED_SERV") * rsPen("PE_BENEFIT_RATE")
''        End If
''        SQLQ = "SELECT * FROM HRP_PA_MASTER WHERE PE_SIN = '" & xSIN & "' "
''        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
''        SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
''        rsPAMaster.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''        If rsPAMaster.EOF Then
''            rsPAMaster.AddNew
''            rsPAMaster("PE_COUNTRY") = rsPen("PE_COUNTRY")
''            rsPAMaster("PE_SIN") = rsPen("PE_SIN")
''            rsPAMaster("PE_EMPNBR") = rsPen("PE_EMPNBR")
''        End If
''        rsPAMaster("PE_SURNAME") = rsPen("PE_SURNAME")
''        rsPAMaster("PE_FNAME") = rsPen("PE_FNAME")
''        rsPAMaster("PE_DIV") = Left(lblCurDiv.Caption, 4) ' rsPen("PE_DIV")
''        rsPAMaster("PE_SECTION") = rsPen("PE_SECTION")
''        rsPAMaster("PE_YEAR_DATE") = xYear
''        rsPAMaster("PE_HRLYSAL") = rsPen("PE_HRLYSAL")
''        If Len(locPayrollID) > 0 Then
''            rsPAMaster("PE_PAYROLL_ID") = locPayrollID
''        End If
''        rsPAMaster("PE_DMD_PENEARN") = xEarnPen
''        '"   (Earned Pension * 9) - 600
''        xTotal = (xEarnPen * 9) - 600
''        If xTotal < 0 Then xTotal = 0
''        rsPAMaster("PE_TOTAL_DBPA") = xTotal
''        rsPAMaster("PE_TOTAL_PA") = xTotal
''        rsPAMaster("PE_LDATE") = Date
''        rsPAMaster("PE_LTIME") = Time$
''        rsPAMaster("PE_LUSER") = glbUserID
''        rsPAMaster.Update
''    End If
''
''End Sub

Private Sub GraniteClubEmpStatusChg()
Dim SQLQ As String
    SQLQ = "UPDATE HREMP SET ED_EMP = 'TERM' WHERE ED_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub WFC_PT_PenCheck(Optional NewHire = "N") 'Ticket #23117 Franks 01/28/2013
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

    'Ticket #23117 Franks 01/28/2013
    xMsg5 = "If this employee qualified for Life, Health or Dental, please notify NGS of the termination. "
    xMsg5 = xMsg5 & "This is needed to end the benefit premiums and qualifications on the NGS side. "
    xMsg5 = xMsg5 & "Info:HR does not send this data to NGS for PT employees."
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    If rsTmpEmp.State <> 0 Then rsTmpEmp.Close
    rsTmpEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmpEmp.EOF Then
        If rsTmpEmp("ED_WORKCOUNTRY") = "U.S.A." And rsTmpEmp("ED_PT") = "PT" And Not rsTmpEmp("ED_DIV") = "1094" Then
            MsgBox xMsg5: Exit Sub
            
            ''check NGS Eligible and Start Date
            'SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1,ER_OTHERDATE2 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
            'rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
            'If Not rsEmpOther.EOF Then
            '    SQLQ = "SELECT * FROM HR_JOB_HISTORY Where JH_EMPNBR = " & glbLEE_ID & " "
            '    SQLQ = SQLQ & "AND NOT JH_CURRENT = 0 "
            '    If rsTmpPos.State <> 0 Then rsTmpPos.Close
            '    rsTmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            '    If Not rsTmpPos.EOF Then
            '        'If rsTmpPos("JH_WHRS") >= 20 Then
            '            MsgBox xMsg5: Exit Sub
            '        'End If
            '    End If
            '    rsTmpPos.Close
            'End If
            'rsEmpOther.Close
        End If
    End If
    rsTmpEmp.Close
End Sub

Private Sub CheckWFCReptAuthExistNew() 'Ticket #29507 Franks 11/30/2016
    If IsWFCReptAuth(glbLEE_ID, "") Then
        glbWFC_IncePlanID = glbLEE_ID
        If glbTermTran Then
             glbWFC_IPPopFormName = "WFCEmpListWithRepTerm"
        Else
            glbWFC_IPPopFormName = "WFCEmpListWithRepTran"
        End If
        frmCheckListView.lblStDate = dlpTermDate.Text
        frmCheckListView.Show 1
    End If
End Sub

Private Sub CheckWFCReptAuthExists() 'Ticket #29343 Franks 10/25/2016
Dim Msg$, DgDef As Variant, Response%
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
Dim SQLQ As String
Dim rsLocEmp As New ADODB.Recordset
Dim xEmpNo

    xFlag1 = False
    
    'for report auth 1,2,3,4
    xEmpNo = glbLEE_ID
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME, ED_SECTION, JH_JOB FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0 "
    SQLQ = SQLQ & "AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
    SQLQ = SQLQ & "AND (JH_REPTAU = " & xEmpNo & " OR JH_REPTAU2 = " & xEmpNo & " OR JH_REPTAU3 = " & xEmpNo & " OR JH_REPTAU4 = " & xEmpNo & ")) "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xFlag1 = True
    End If

    If xFlag1 Then
        glbNewRept = ""
        locWFCEmpID = glbLEE_ID
        xMsg = "This employee is a Reporting Authority to one or more employees." & Chr(10) & Chr(10)
        'xMsg = xMsg & "Is there another employee who is resuming this responsibility?"
        'Ticket #29343 Franks 10/24/2016
        xMsg = xMsg & "Is there another employee who is assuming this responsibility?"
        Response% = MsgBox(xMsg, vbQuestion + vbYesNo, "Reporting Authority")
        If Response% = IDYES Then
            glbNewRept = ""
            frmMsgRA.lblWFCTermNo = glbLEE_ID
            frmMsgRA.Show 1
        End If
        
        If Len(glbNewRept) > 0 Then
            If gsEMAIL_ONPOSITION Then
                xWFCPosChgEmailBody = "Teammate #" & xEmpNo & " " & lblEEName.Caption & " was terminated on " & dlpTermDate.Text & ". "
                xWFCPosChgEmailBody = xWFCPosChgEmailBody & "The Interim or New Reporting Authority # " & glbNewRept & " " & GetEmpData(glbNewRept, "ED_SURNAME") & "," & GetEmpData(glbNewRept, "ED_FNAME") & " "
                xWFCPosChgEmailBody = xWFCPosChgEmailBody & "with an Effective date of " & DateAdd("d", 1, CVDate(dlpTermDate.Text)) & " "
                xWFCPosChgEmailBody = xWFCPosChgEmailBody & "has been added to the following list of employees: " & Chr(10)
            End If
        End If
        
        Call PubReptPosEmpListByEmp(glbLEE_ID) 'Ticket #29507 Franks 11/29/2016
        
        Call ShowPosEmpListByEmp("Report To " & lblEEName.Caption & " Employee list")
        
        Do While Not rsLocEmp.EOF
            'Call WFCPosReptsUpd(xEmpNo, xOldReptEmpNo, xNewReptEmpNo, xEffDate) 'Ticket #29343 Franks 10/24/2016
            Call WFCPosReptsUpd(rsLocEmp("ED_EMPNBR"), xEmpNo, glbNewRept, DateAdd("d", 1, CVDate(dlpTermDate.Text)))
            If Len(glbNewRept) > 0 Then
                If gsEMAIL_ONPOSITION Then
                    xWFCPosChgEmailBody = xWFCPosChgEmailBody & GetTABLDesc("EDSE", rsLocEmp("ED_SECTION")) & " - Employee #" & rsLocEmp("ED_EMPNBR") & "/" & rsLocEmp("ED_SURNAME") & "," & rsLocEmp("ED_FNAME") & " - "
                    xWFCPosChgEmailBody = xWFCPosChgEmailBody & "Position " & rsLocEmp("JH_JOB") & "/" & getPosDesc(rsLocEmp("JH_JOB")) & Chr(10)
                    xIsWFCPosChgEmail = True
                End If
            End If
            rsLocEmp.MoveNext
        Loop

    End If
    
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
End Sub

Private Sub CheckReptAuthExists()
Dim Msg$, DgDef As Variant, Response%
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
Dim SQLQ As String

    xFlag1 = False
    
    'check if this employee is a Reporting Authority
    If IsReportAuth(glbLEE_ID) Then
        xFlag1 = True
    End If

    If xFlag1 Then
        glbNewRept = ""
        xMsg = "This employee is a Reporting Authority to one or more employees." & Chr(10) & Chr(10)
        'xMsg = xMsg & "Is there another employee who is resuming this responsibility?"
        'Ticket #29343 Franks 10/24/2016
        xMsg = xMsg & "Is there another employee who is assuming this responsibility?"
        Response% = MsgBox(xMsg, vbQuestion + vbYesNo, "Reporting Authority")
        If Response% = IDYES Then
            glbNewRept = ""
            
            frmMsgRA.Show 1
            
        End If
        
        If glbNewRept = "" Then
            'This employee will be automatically removed from their Reporting Authorities.
            'Reporting Authority 1
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 2
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU2 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU2 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 3
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU3 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU3 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 4
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = NULL WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU4 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
        Else
            'New employee will be assigned as a Reporting Authorities.
            'Reporting Authority 1
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = " & glbNewRept & " WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 2
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU2 = " & glbNewRept & " WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU2 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 3
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU3 = " & glbNewRept & " WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU3 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
            'Reporting Authority 4
            SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = " & glbNewRept & " WHERE NOT (JH_CURRENT = 0) AND JH_REPTAU4 = " & glbLEE_ID & " "
            gdbAdoIhr001.Execute SQLQ
        End If

    End If
End Sub


''Private Sub WFC_UptPenDate4WithDOT(xEmpNo, xDOT) 'Ticket #23948 Frank 06/24/2013
'''Termination -
'''"   Always update Pension Date 4 with the Date of Termination. Remove the check for Eligible for Pension.
''Dim rsOther As New ADODB.Recordset
''Dim SQLQ As String
''    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
''    If rsOther.State <> 0 Then rsOther.Close
''    rsOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''    If Not rsOther.EOF Then
''        rsOther("ER_PENSIONDATE4") = xDOT
''        rsOther.Update
''    End If
''    rsOther.Close
''End Sub

'Ticket #23247 Franks 07/22/2013
Private Sub chkAllDates_Click()
Dim SQLQ As String
Dim xID As Long
    If IsDate(dlpEndDate.Text) Then
        If Not Data2.Recordset.EOF Then
            xID = Data2.Recordset("BM_BENE_ID")
            If chkAllDates.Value Then 'checked
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(dlpEndDate.Text) & " WHERE BM_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            Else 'unchecked
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = Null WHERE BM_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                SQLQ = SQLQ & "AND NOT (BM_BENE_ID = " & xID & ") "
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            End If
        End If
    End If
End Sub
'Ticket #23247 Franks 07/22/2013
Private Sub UptData2fromDOT()
Dim SQLQ As String
Dim xID As Long
If glbWFC And dlpDOther2.Visible Then
    If IsDate(dlpDOther2.Text) Then  '(dlpTermDate.Text) Then
        'Ticket #24167 - Getting an error 91. When there is no xNGSStart (assigned in WFCOther2Screen function) the
        'Data2 is not set. So when in this fuction is called the 'Not Data2.Recordset.EOF' gives an error.
        'So I am checking if Data2.RecordSource = "" to avoid the error. I am not too sure of this logic so I am
        'just adding this to avoid error.
        If Data2.RecordSource <> "" Then
            If Not Data2.Recordset.EOF Then
                xID = Data2.Recordset("BM_BENE_ID")
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(dlpDOther2.Text) & " WHERE BM_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            End If
        End If
    End If
End If
End Sub

Private Sub dlpEndDate_Change() 'Ticket #23247 Franks 07/22/2013
xLocID = 0
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    xLocID = Data2.Recordset("BM_BENE_ID")
End If
End Sub

Private Sub dlpEndDate_LostFocus() 'Ticket #23247 Franks 07/22/2013
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    If Data2.Recordset("BM_PCE") = 1 Then 'employee %
        If IsNull(Data2.Recordset("BM_ENDDATE")) Then
            If IsDate(dlpEndDate.Text) Then
                MsgBox "Cannot enter END DATE on 100% paid employee benefits."
                dlpEndDate.Text = ""
                Exit Sub
            End If
        End If
    End If
End If
Call WFCUpdate_Value
End Sub

Private Sub WFCBenListScreen(xEmpNo) 'Ticket #23247 Franks 07/22/2013
Dim rsLEmp As New ADODB.Recordset
Dim rslBen As New ADODB.Recordset
Dim SQLQ As String
    txtComments.Height = 700
    frmWFCBenList.Top = 7000
    frmWFCBenList.Left = 130
    frmWFCBenList.Width = 10575
    frmWFCBenList.Height = 2055 - 330 ' 2295 '2175
    chkAllDates.Caption = "All Dates"
    chkAllDates.Value = False
    Call WFCUpdateBenefitGroup(xEmpNo)
    
    Data2.ConnectionString = glbAdoIHRDBW
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    Data2.RecordSource = SQLQ
    Data2.Refresh
                        
    frmWFCBenList.Visible = True
End Sub

''Private Sub WFCUpdateBenefitGroup(xEmpNo) 'Ticket #23247 Franks 07/22/2013
''Dim rsBGMST As New ADODB.Recordset
''Dim rsBGTMP As New ADODB.Recordset
''Dim rsBGEE As New ADODB.Recordset
''Dim rsTABL As New ADODB.Recordset
''Dim SQLQ As String
''Dim BelongOldGroup As Boolean
''    gdbAdoIhr001W.BeginTrans
''    gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
''    gdbAdoIhr001W.CommitTrans
''
''    gdbAdoIhr001W.BeginTrans
''    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
''    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
''
''    SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
''    'SQLQ = SQLQ & "AND BF_PCC = 1 " 'Paid Benefits only
''    SQLQ = SQLQ & "ORDER BY BF_BCODE, BF_EDATE "
''
''    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
''
''    Do While Not rsBGMST.EOF
''        rsBGTMP.AddNew
''        rsBGTMP("BM_COMPNO") = "001"
''        rsBGTMP("BM_BENEFIT_GROUP") = rsBGMST("BF_GROUP")
''        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
''        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
''        rsBGTMP("BM_ENDDATE") = rsBGMST("BF_CEASEDATE") 'New
''        rsBGTMP("BM_CHECK") = 1
''        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
''        rsBGTMP("BM_AMT") = rsBGMST("BF_AMT")
''        rsBGTMP("BM_PPAMT") = rsBGMST("BF_PPAMT")
''        rsBGTMP("BM_UNITCOST") = rsBGMST("BF_UNITCOST")
''        rsBGTMP("BM_PCE") = rsBGMST("BF_PCE")
''        rsBGTMP("BM_PCC") = rsBGMST("BF_PCC")
''        rsBGTMP("BM_ECOST") = rsBGMST("BF_ECOST")
''        rsBGTMP("BM_CCOST") = rsBGMST("BF_CCOST")
''        rsBGTMP("BM_TCOST") = rsBGMST("BF_TCOST")
''        rsBGTMP("BM_MAXDOL") = rsBGMST("BF_MAXDOL")
''        rsBGTMP("BM_PREMIUM") = rsBGMST("BF_PREMIUM")
''        rsBGTMP("BM_PER") = rsBGMST("BF_PER")
''        rsBGTMP("BM_MTHCCOST") = rsBGMST("BF_MTHCCOST")
''        rsBGTMP("BM_MTHECOST") = rsBGMST("BF_MTHECOST")
''        rsBGTMP("BM_TAXBEN") = rsBGMST("BF_TAXBEN")
''        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BF_SALARYDEPENDANT")
''        rsBGTMP("BM_MINIMUM") = rsBGMST("BF_MINIMUM")
''        rsBGTMP("BM_FACTOR") = rsBGMST("BF_FACTOR")
''        rsBGTMP("BM_ROUND") = rsBGMST("BF_ROUND")
''        rsBGTMP("BM_MAXIMUM") = rsBGMST("BF_MAXIMUM")
''        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BF_NEXTNEAREST")
''        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BF_TAXAMOUNT")
''        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BF_WAITPERIOD")
''        rsBGTMP("BM_DWM") = rsBGMST("BF_DWM")
''        rsBGTMP("BM_PERORDOLL") = rsBGMST("BF_PERORDOLL")
''        rsBGTMP("BM_POLICY") = rsBGMST("BF_POLICY")
''        rsBGTMP("BM_RATELEVEL") = rsBGMST("BF_RATELEVEL")
''        rsBGTMP("BM_COMMENTS") = rsBGMST("BF_COMMENTS")
''        rsBGTMP("BM_PTAX") = rsBGMST("BF_PTAX")
''        rsBGTMP("BM_ACTION") = "Add"
''        rsBGTMP("BM_WRKEMP") = glbUserID
''
''        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
''        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
''        If Not rsTABL.EOF Then
''            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
''        End If
''        rsTABL.Close
''        rsBGTMP.Update
''        rsBGMST.MoveNext
''    Loop
''    rsBGTMP.Close
''    rsBGMST.Close
''    gdbAdoIhr001W.CommitTrans
''    Call Pause(1)
''
''End Sub

Private Sub WFCUpdate_Value() 'Ticket #23247 Franks 07/22/2013
Dim SQLQ As String
Dim xID As Long
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    
    xID = Data2.Recordset("BM_BENE_ID")
    Data2.Refresh
    'xID = xLocID
    If Not IsEmpty(xLocID) Then
        If xLocID > 0 Then
            xID = xLocID
        End If
    End If
    SQLQ = "BM_BENE_ID = " & xID 'xLocID
    Data2.Recordset.Find SQLQ
    
    If IsDate(dlpEndDate.Text) Then
        If Year(dlpEndDate.Text) > 1900 And Year(dlpEndDate.Text) < 2050 Then
            Data2.Recordset("BM_ENDDATE") = dlpEndDate.Text
        Else
            Data2.Recordset("BM_ENDDATE") = Null
        End If
    Else
        Data2.Recordset("BM_ENDDATE") = Null
    End If
    Data2.Recordset.Update
    Data2.Refresh
    DoEvents
    SQLQ = "BM_BENE_ID = " & xID
    Data2.Recordset.Find SQLQ
End If
End Sub

'Ticket #23247 Franks 07/22/2013
Private Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
        If IsNull(Data2.Recordset("BM_ENDDATE")) Then
            dlpEndDate.Text = ""
        Else
            dlpEndDate.Text = Data2.Recordset("BM_ENDDATE")
        End If
    End If
End Sub

Private Sub WFC_LastDayUpt(xEmpNo) 'Ticket #26308 Franks 11/27/2014
Dim SQLQ
    'update Last Day
    If IsDate(dlpLastDate.Text) Then
        SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(dlpLastDate.Text) & " "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
        gdbAdoIhr001.Execute SQLQ
        Call WFCAUDITBENF_NGSEnd(xEmpNo, False, , "Y", dlpLastDate.Text)
    End If
End Sub

Private Sub NonWFC_BenEndDateUp(xEmpNo) '09/02/2015 non WFC Diff Benefit End Date
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim rsEmpBN As New ADODB.Recordset
Dim xTemp
Dim xDate1, xDate2
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBN.EOF
        SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & rsBN("BM_BCODE") & "' "
        If Not IsNull(rsBN("BM_EDATE")) Then SQLQ = SQLQ & "AND BF_EDATE = " & Date_SQL(rsBN("BM_EDATE")) & " "
        If rsEmpBN.State <> 0 Then rsEmpBN.Close
        rsEmpBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpBN.EOF Then
            If IsNull(rsEmpBN("BF_CEASEDATE")) Then xDate1 = CVDate("01/01/1900") Else xDate1 = CVDate(rsEmpBN("BF_CEASEDATE"))
            If IsNull(rsBN("BM_ENDDATE")) Then xDate2 = CVDate("01/01/1900") Else xDate2 = CVDate(rsBN("BM_ENDDATE"))
            rsEmpBN("BF_CEASEDATE") = rsBN("BM_ENDDATE")
            rsEmpBN.Update
            If Not xDate1 = xDate2 Then 'BF_CEASEDATE was changed
                If xDate2 > CVDate("01/01/1900") Then
                    'update hraudit - begin
                    Call NonWFCAUDITBENF_End(xEmpNo, False, rsEmpBN) 'Call WFCAUDITBENF_NGSEnd(xEmpNo, False, rsEmpBN)
                    'update hraudit - end
                End If
            End If
        End If
        rsEmpBN.Close
        rsBN.MoveNext
    Loop
    rsBN.Close

End Sub

Private Sub WFC_NGSBenEndDateUpt(xEmpNo) 'Ticket #23247 Franks 07/22/2013
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim rsEmpBN As New ADODB.Recordset
Dim xTemp
Dim xDate1, xDate2
    'update Last Day
    If IsDate(dlpLastDate.Text) Then
        SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(dlpLastDate.Text) & " "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
        gdbAdoIhr001.Execute SQLQ
        Call WFCAUDITBENF_NGSEnd(xEmpNo, False, , "Y", dlpLastDate.Text)
    End If
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBN.EOF
        SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & rsBN("BM_BCODE") & "' "
        If Not IsNull(rsBN("BM_EDATE")) Then SQLQ = SQLQ & "AND BF_EDATE = " & Date_SQL(rsBN("BM_EDATE")) & " "
        If rsEmpBN.State <> 0 Then rsEmpBN.Close
        rsEmpBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpBN.EOF Then
            If IsNull(rsEmpBN("BF_CEASEDATE")) Then xDate1 = CVDate("01/01/1900") Else xDate1 = CVDate(rsEmpBN("BF_CEASEDATE"))
            If IsNull(rsBN("BM_ENDDATE")) Then xDate2 = CVDate("01/01/1900") Else xDate2 = CVDate(rsBN("BM_ENDDATE"))
            rsEmpBN("BF_CEASEDATE") = rsBN("BM_ENDDATE")
            rsEmpBN.Update
            If Not xDate1 = xDate2 Then 'BF_CEASEDATE was changed
                If xDate2 > CVDate("01/01/1900") Then
                    'update hraudit - begin
                    Call WFCAUDITBENF_NGSEnd(xEmpNo, False, rsEmpBN)
                    'update hraudit - end
                End If
            End If
        End If
        rsEmpBN.Close
        rsBN.MoveNext
    Loop
    rsBN.Close
End Sub

'Ticket #23247 Franks 07/22/2013
''Private Function WFCAUDITBENF_NGSEnd(xEmpNo, xlocNewRec As Boolean, Optional rslBen As ADODB.Recordset, Optional xIsWorkDay = "N", Optional xLastDate)
''Dim rsEmp As New ADODB.Recordset
''Dim rsTA As New ADODB.Recordset
''Dim rsTB As New ADODB.Recordset
''Dim xADD As Boolean, xPT As String, xDiv As String
''Dim strFields As String
''Dim ACTX
''Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
''Dim xTermSEQ
''Dim SQLQ As String
''
'''''On Error GoTo AUDIT_ERR
''WFCAUDITBENF_NGSEnd = False
''
''If xlocNewRec Then
''    ACTX = "A"
''Else
''    ACTX = "M"
''End If
''
''xTermSEQ = 0
''If xTermSEQ = 0 Then
''    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
''Else
''    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
''    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
''End If
''rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
''
''If Not rsTB.EOF Then
''    If IsNull(rsTB("ED_PT")) Then
''        xPT = ""
''    Else
''        xPT = rsTB("ED_PT")
''    End If
''    If IsNull(rsTB("ED_DIV")) Then
''        xDiv = ""
''    Else
''        xDiv = rsTB("ED_DIV")
''    End If
''Else
''    xPT = ""
''    xDiv = ""
''End If
'''strfields added by Bryan 02/Dec/05 Ticket#9899
''strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
''strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
''strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
''strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
''strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
''rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''
''xADD = False
''
''If xIsWorkDay = "N" Then
''    NBCode = ""
''    NPPAMT = ""
''    NMTHCOMP = ""
''    NMTHEMP = ""
''    NBAMT = ""
''    NPPE = ""
''    NPCC = ""
''    NMAXDOL = ""
''    NEDate = ""
''    NCOVER = ""
''    NTCOST = ""
''    NBCode = rslBen("BF_BCODE")
''    If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
''    ''If Not IsNull(rslBen("BF_PPAMT")) Then NPPAMT = rslBen("BF_PPAMT")
''    ''If Not IsNull(rslBen("BF_MTHCCOST")) Then NMTHCOMP = rslBen("BF_MTHCCOST")
''    ''If Not IsNull(rslBen("BF_MTHECOST")) Then NMTHEMP = rslBen("BF_MTHECOST")
''    ''If Not IsNull(rslBen("BF_AMT")) Then NBAMT = rslBen("BF_AMT")
''    ''If Not IsNull(rslBen("BF_PCC")) Then NPCC = rslBen("BF_PCC")
''    ''If Not IsNull(rslBen("BF_PCE")) Then NPPE = rslBen("BF_PCE")
''    ''If Not IsNull(rslBen("BF_MAXDOL")) Then NMAXDOL = rslBen("BF_MAXDOL")
''    ''If Not IsNull(rslBen("BF_COVER")) Then NCOVER = rslBen("BF_COVER")
''    ''If Not IsNull(rslBen("BF_TCOST")) Then NTCOST = rslBen("BF_TCOST")
''    ''
''    ''If OBCode <> NBCode Then GoTo MODUPD
''    '''If OPPE <> NPPE Or OPCC <> NPCC Then GoTo MODUPD
''    ''If OPPAMT <> NPPAMT Or OMAXDOL <> NMAXDOL Then GoTo MODUPD
''    '''If OMTHCOMP <> NMTHCOMP Or OMTHEMP <> NMTHEMP Then GoTo MODUPD
''    ''If OBAMT <> NBAMT Then GoTo MODUPD
''    ''If OEDate <> NEDate Then GoTo MODUPD
''End If
''
'''GoTo MODNOUPD
''
'''BF_CEASEDATE was changed
''MODUPD:
''
''rsTA.AddNew
''rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
''rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
''rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
''rsTA("AU_NEWEMP") = "N"
''rsTA("AU_PTUPL") = xPT
''rsTA("AU_DIVUPL") = xDiv
''
''If xIsWorkDay = "N" Then
''    rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
''    rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
''    'If OMTHCOMP <> NMTHCOMP Then rsTA("AU_MTHCCOST") = NMTHCOMP
''    'If OMTHEMP <> NMTHEMP Then rsTA("AU_MTHECOST") = NMTHEMP
''    'If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN
''    'If OCOVER <> NCOVER Then rsTA("AU_COVER") = NCOVER
''    'If OTCOST <> NTCOST Then rsTA("AU_TCOST") = NTCOST
''    'If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
''    'If OPPE <> NPPE Then rsTA("AU_PCE") = NPPE
''    'If OPCC <> NPCC Then rsTA("AU_PCC") = NPCC
''    'If OPPAMT <> NPPAMT Then
''    '    rsTA("AU_PPAMT") = NPPAMT
''    '    If IsNumeric(OPPAMT) Then rsTA("AU_OLDPPMT") = Val(OPPAMT)
''    'End If
''    'If OMAXDOL <> NMAXDOL Then rsTA("AU_MAXDOL") = NMAXDOL
''    'If OEDate <> NEDate Then
''    '  If IsDate(NEDate) Then
''    '      rsTA("AU_EDATE") = CVDate(NEDate)
''    '  End If
''    'End If
''    'If OPER <> txtPer Then rsTA("AU_PER") = txtPer
''    'If OBAMT <> NBAMT Then rsTA("AU_BAMT") = NBAMT
''    'If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)
''    rsTA("AU_LDATE") = Date
''    If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
''        If CVDate(NEDate) > CVDate(Date) Then
''            rsTA("AU_LDATE") = CVDate(NEDate)
''        End If
''    End If
''End If
''If xIsWorkDay = "Y" Then
''    rsTA("AU_LDAY") = xLastDate
''    rsTA("AU_LDATE") = Date
''End If
''If xTermSEQ = 0 Then
''    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
''Else
''    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
''    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
''End If
''rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
''If Not rsEmp.EOF Then
''    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
''End If
''rsEmp.Close
''rsTA("AU_COMPNO") = "001"
''rsTA("AU_EMPNBR") = xEmpNo
''rsTA("AU_LUSER") = glbUserID
''rsTA("AU_LTIME") = Time$
''rsTA("AU_UPLOAD") = "N"
''rsTA("AU_TYPE") = ACTX
''rsTA.Update
''rsTA.Close
''
''MODNOUPD:
''WFCAUDITBENF_NGSEnd = True
''Exit Function
''AUDIT_ERR:
''
''End Function

''Private Sub UptLUserLDateLTime() 'Ticket #24355 Franks 09/17/2013
''Dim SQLQ As String
''    'update Term_HREMP
''    SQLQ = "UPDATE Term_HREMP SET ED_LDATE = " & Date_SQL(Date) & ", "
''    SQLQ = SQLQ & "ED_LUSER = '" & glbUserID & "', "
''    SQLQ = SQLQ & "ED_LTIME = '" & Time$ & "' "
''    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
''    gdbAdoIhr001X.Execute SQLQ
''
''    'update Term_HRTRMEMP
''    SQLQ = "UPDATE Term_HRTRMEMP SET Term_LDATE = " & Date_SQL(Date) & ", "
''    SQLQ = SQLQ & "Term_LUSER = '" & glbUserID & "', "
''    SQLQ = SQLQ & "Term_LTIME = '" & Time$ & "' "
''    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
''    gdbAdoIhr001X.Execute SQLQ
''End Sub

Private Sub WFCNGSEndDateForTransferOut() 'Ticket #24767 Franks 12/11/2013
'"   If employee is transferring from a NGS plant to a non-NGS plant
'the company paid benefits should have an end date equal to the Transfer Out date
If glbWFC Then
    If dlpDOther2.Visible Then
        If IsDate(dlpTermDate.Text) Then
            If Not glbTermTran Then 'transfer out
                If Len(comDIV.Text) > 0 Then
                    If Not IsWFCNGSDiv(Left(comDIV.Text, 4)) Then
                        If Len(dlpDOther2.Text) = 0 Then
                            dlpDOther2.Text = dlpTermDate.Text
                            Call UptData2fromDOT
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub FamilyDay2ndTermScreen(x2ndEmpID)
    Call FoundLocEEFind(x2ndEmpID)
    Call cll_EEFind(Me)
End Sub

Private Sub WFCDivTranSamePlantScreen()
    lblCurUnion.Caption = GetEmpData(glbLEE_ID, "ED_ORG")
    lblCurDiv.Caption = GetEmpData(glbLEE_ID, "ED_DIV")
    '--------- new - begin
    cmdPrintSelected.Visible = False
    panTermRpts.Visible = False
    lblImport.Visible = False
    cmdImport.Visible = False
    imgNoSec.Visible = False
    imgSec.Visible = False
    chkSum.Visible = False
    chkRehire.Visible = False
    lblTitle(5).Visible = False 'Benefit End Date
    dlpBenCeaseDate.Visible = False
    lblTitle(2).Visible = False 'comment
    txtComments.Visible = False
End Sub

Private Sub WFCNormalTranOutScreen()
        lblTitle(10).Caption = "Transfer to " & lStr("Union")
        lblTitle(10).Visible = True
        clpCode(3).Visible = True
        'lblCurUnion.Visible = True
        lblTitle(10).Top = 1295 'lblTitle(4).Top
        lblTitle(10).Left = lblTitle(5).Left
        lblCurUnion.Left = 6000
        lblCurUnion.Top = lblTitle(10).Top
        clpCode(3).Top = 1295 ' - Union
        clpCode(3).Left = clpCode(2).Left
        lblTitle(4).Top = 1630 'Employee #
        lblEMPNo.Top = 1630
        lblTitle(5).Top = 1900 'Benefit End Date
        dlpBenCeaseDate.Top = 1900
        
        'lblCurUnion.Caption = "Current " & lStr("Union") & ": " & GetEmpData(glbLEE_ID, "ED_ORG")
        lblCurUnion.Caption = GetEmpData(glbLEE_ID, "ED_ORG")
        lblCurDiv.Caption = GetEmpData(glbLEE_ID, "ED_DIV")
        
        'Ticket #27827 Franks 12/14/2015
        clpCode(3).Width = 1200
        lblWFCUnion.Top = lblCurUnion.Top
        lblWFCUnion.Left = clpCode(3).Left + 1300 '2500
        lblWFCUnion.Visible = True
End Sub

Private Sub WFCScreenSetup() 'Ticket #26308 Franks 11/27/2014
    cmdEmail.Visible = False
    lblTitle(3).Caption = lStr("Transfer to Division")
    glbWFCEmailTest = False
    'Ticket #16748 - begin
    'lblTitle(6).Visible = True
    'cboStatFlag3.Visible = True
    'If glbTermTran Then
    '    lblTitle(6).Top = lblTitle(5).Top
    '    cboStatFlag3.Top = dlpBenCeaseDate.Top
    'End If
    'Ticket #16748 - end
    
    txtComments.Height = 700
    
    lblTitle(5).Visible = True
    dlpBenCeaseDate.Visible = True
    If glbTermTran Then
        lblTitle(5).Top = 960
        dlpBenCeaseDate.Top = 960
    End If
    
    'Ticket #15248 Termination Cause - Begin
    If glbTermTran Then
        lblTitle(0).Top = 0
        dlpTermDate.Top = 0
        lblTitle(1).Top = 330
        clpCode(1).Top = 330
        lblTitle(7).Top = 640
        clpCode(2).Top = 640
        lblTitle(7).Visible = True
        clpCode(2).Visible = True
        lblTitle(7).FontBold = True 'Ticket #27820 Franks 11/25/2015
        
        'Ticket #26308 Franks 11/27/2014 - begin
        dlpTermDate.ShowDescription = False
        dlpTermDate.Width = 1500
        frmLastDay.Top = dlpTermDate.Top
        frmLastDay.Visible = True
        frmLastDay.BorderStyle = 0
        'Ticket #26308 Franks 11/27/2014 - end
    End If
    'Ticket #15248 Termination Cause - End

    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'For Frank Test with WFC database
        cmdFrankTest.Visible = True
    End If
        
End Sub

Private Sub ShowPosEmpListByEmp(xTitle)

    Me.vbxCrystal3.ReportFileName = glbIHRREPORTS & "RZEmpList4.rpt" '"RZEmpList3.rpt"
    Me.vbxCrystal3.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal3.Formulas(0) = "rTitle='" & xTitle & " '"  '"rTitle='" & lblReptAuth(0).Caption & " information'"
    Me.vbxCrystal3.Connect = RptODBC_SQL
    'window title if appropriate
    Me.vbxCrystal3.WindowTitle = xTitle 'lblReptAuth(0).Caption & " Employee Position Information"
    Me.vbxCrystal3.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal3.Action = 1
    vbxCrystal3.Reset

End Sub
