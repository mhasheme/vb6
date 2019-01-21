VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRPosMaster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Positions / Skills / Evaluation Factors Report"
   ClientHeight    =   6645
   ClientLeft      =   210
   ClientTop       =   885
   ClientWidth     =   10785
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
   ScaleHeight     =   6645
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.TextBox medLevel 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
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
      Left            =   1880
      TabIndex        =   11
      Tag             =   "00-Level"
      Top             =   3330
      Visible         =   0   'False
      Width           =   4005
   End
   Begin INFOHR_Controls.CodeLookup clpNationalClass 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1350
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   6
   End
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Skill Code"
      Top             =   2670
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSK"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1020
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBST"
      MaxLength       =   6
   End
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   2010
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   690
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin VB.Frame frePaper 
      Caption         =   "Paper Size"
      Height          =   855
      Left            =   3720
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      Begin Threed.SSOption optPaper 
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "Select Evaluation Reports"
         Top             =   540
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Legal"
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
      Begin Threed.SSOption optPaper 
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "Select Evaluation Reports"
         Top             =   270
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Letter"
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Third Level of Records Grouping"
      Top             =   6135
      Width           =   2115
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Second Level of Records Grouping"
      Top             =   5820
      Width           =   2115
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "First Level of Records Grouping"
      Top             =   5505
      Width           =   2115
   End
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   0
      Left            =   1860
      TabIndex        =   12
      Tag             =   "Select Detail Reports"
      Top             =   3810
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Detail "
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "Select Reports Showing Totals Only"
      Top             =   4050
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   2
      Left            =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "Select Skill Reports"
      Top             =   4290
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Skills"
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   3
      Left            =   1860
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "Select Evaluation Reports"
      Top             =   4530
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Evaluation"
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   4
      Left            =   1860
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "Select Evaluation Reports"
      Top             =   4770
      Visible         =   0   'False
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Position Grid "
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
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   10
      Tag             =   "00-Enter Evaluation Factor Code"
      Top             =   3000
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBFC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6840
      Top             =   6600
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
   Begin INFOHR_Controls.CodeLookup clpGrid 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   2340
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGD"
   End
   Begin VB.CheckBox chkIncPos 
      Caption         =   "Exclude Inactive Positions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin INFOHR_Controls.CodeLookup clpcode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Machine #"
      Top             =   2340
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   12
   End
   Begin INFOHR_Controls.CodeLookup clpJobMaster 
      Height          =   285
      Left            =   7500
      TabIndex        =   1
      Tag             =   "01-Job code"
      Top             =   360
      Visible         =   0   'False
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   13
   End
   Begin VB.Label lblJob 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
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
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   675
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
      Left            =   120
      TabIndex        =   40
      Top             =   2370
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLevel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
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
      Top             =   3330
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Category"
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
      Top             =   2370
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Skills"
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
      TabIndex        =   36
      Top             =   2700
      Width           =   360
   End
   Begin VB.Label lblGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   6165
      Width           =   825
   End
   Begin VB.Label lblGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      TabIndex        =   34
      Top             =   5850
      Width           =   1065
   End
   Begin VB.Label lblGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      TabIndex        =   33
      Top             =   5535
      Width           =   1065
   End
   Begin VB.Label lblRepGrp 
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
      TabIndex        =   32
      Top             =   5295
      Width           =   1575
   End
   Begin VB.Label lblSelCri 
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
      TabIndex        =   31
      Top             =   120
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
      Index           =   8
      Left            =   120
      TabIndex        =   30
      Top             =   3810
      Width           =   1065
   End
   Begin VB.Label lblEval 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluation"
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
      TabIndex        =   29
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label lblPosTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reports To"
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
      TabIndex        =   28
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label lblProv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N.O.C. Code"
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
      TabIndex        =   27
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Status"
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
      TabIndex        =   26
      Top             =   1020
      Width           =   1050
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
      TabIndex        =   25
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   690
      Width           =   1035
   End
   Begin VB.Label lblPosTitle 
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
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmRPosMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub clpCode_LostFocus(Index As Integer)
On Error Resume Next:
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Position Master Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
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
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub



Private Sub CodeLookup1_Change()

End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim vPosGroup As String
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(3).Clear
    If Not glbSyndesis Then
        vPosGroup = "Position Group"
    Else
        vPosGroup = "Position Grade"
    End If
    comGroup(1).AddItem vPosGroup '"Position Group"
    If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #15590
        comGroup(1).AddItem "Income Code"
    Else
        comGroup(1).AddItem "Position Status"
    End If
    comGroup(1).AddItem lStr("Union")
    If optGrouping(3) Then
        comGroup(1).AddItem "Level"
    End If
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    comGroup(2).AddItem lStr("Union")
    comGroup(2).AddItem "Reports To"
    comGroup(2).AddItem vPosGroup '"Position Group"
    If optGrouping(3) Then
        comGroup(2).AddItem "Level"
    End If
    comGroup(2).AddItem "(none)"
    comGroup(2).ListIndex = 0
    comGroup(3).AddItem "Position Description"
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False
End Sub



Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpcode(intIdx%).Text) > 0 Then
    If intIdx% = 1 Then
        strCd$ = "HRJOB.JB_GRPCD"
    ElseIf intIdx% = 2 Then
        If glbMultiGrid Then
            strCd$ = "HRJOB_GRADE.JB_ORG"
        Else
            strCd$ = "HRJOB.JB_ORG"
        End If
    ElseIf intIdx% = 3 Then
        strCd$ = "HRJOB.JB_STATUS"
    End If
       CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpcode(intIdx%).Text, ",", "','") & "'])"
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
Private Sub Cri_Level()
Dim EECri As String, OneSet%, X%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False

If Len(medLevel.Text) > 0 Then
    OneSet% = OneSet% + 1
End If

If OneSet% = 0 Then Exit Sub

strCa$ = "HRJOBEVL.JE_LEVEL"

EECri = "({" & strCa$ & "} in  ['" & Replace(medLevel.Text, ",", "','") & "'])"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub
Private Sub Cri_Evaluation()
'Laura
Dim EECri As String, OneSet%, X%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False

For X% = 5 To 5 ' 11
    If Len(clpcode(X%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then Exit Sub

strCa$ = "HRJOBEVL.JE_FACTOR"

For X% = 5 To 5
  EECri = "({" & strCa$ & "} in  ['" & Replace(clpcode(X%).Text, ",", "','") & "'])"
Next X%

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True



End Sub

Private Sub Cri_JobRep()
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim X%

For X% = 0 To 1
    Select Case X%
        Case 0: fld$ = "HRJOB.JB_CODE"
        Case 1: fld$ = "HRJOB.JB_REPTAU"
    End Select


If Len(clpJob(X%).Text) > 0 Then
    CodeCri = "({" & fld$ & "} = '" & clpJob(X%).Text & "')"
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

Next X%

End Sub
Private Sub Cri_Grid()
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim X%

If Len(clpGrid.Text) > 0 Then
    CodeCri = "({HRJOB_GRADE.JB_GRID} = '" & clpGrid.Text & "')"
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
Private Sub Cri_NOC()
Dim CodeCri As String, fld$
Dim X%


fld$ = "HR_OCCUPATION_CLASS.OC_CODE"


If Len(clpNationalClass) > 0 Then
    CodeCri = "({" & fld$ & "} = '" & clpNationalClass.Text & "')"
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

Private Function Cri_SetAll()
Dim X%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
Call Cri_Code(1)
Call Cri_Code(2)
Call Cri_Code(3)
Call Cri_Skills
If optGrouping(3) Then
    Call Cri_Evaluation
    Call Cri_Level
End If
Call Cri_JobRep
Call Cri_Grid

If glbWFC Then 'Ticket #27531 Franks 09/15/2015
    If clpJobMaster.Visible Then
        Call Cri_JobMaster
    End If
End If

X% = mod_Upd_Pos_Totals(True)

' report name
If optGrouping(0).Value = True Then
    If glbMultiGrid Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzMGmst.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzposmst.rpt"
    End If
    Me.vbxCrystal.Connect = RptODBC_SQL
ElseIf optGrouping(1).Value = True Then
    If glbMultiGrid Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzMGmts.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzposmts.rpt"
    End If
    Me.vbxCrystal.Connect = RptODBC_SQL
ElseIf optGrouping(2).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzposskl.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
ElseIf optGrouping(3).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzposevl.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
ElseIf optGrouping(4).Value = True Then
    If optPaper(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2191-1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2191-2.rpt"
    End If
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 3
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next X%
End If
' set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
'Modified by Franks on Feb 12,2002 for error when the user's union set to "-NON"
'Because no HREMP in these reports
'If glbNoNONE Then
'    If Len(glbstrSelCri) > 0 Then
'        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND ({HREMP.ED_ORG }<> 'NONE')"
'    Else
'        Me.vbxCrystal.SelectionFormula = "({HREMP.ED_ORG } <> 'NONE')"
'    End If
'Else

'Ticket #13244 Frank Jun 20, 2007
If chkIncPos Then
    If Len(glbstrSelCri) = 0 Then
        glbstrSelCri = " NOT ({HRJOB.JB_STATUS} = 'INAC' or Left({HRJOB.JB_DESCR}, 2) = 'Z ')"
    Else
        glbstrSelCri = glbstrSelCri & " AND NOT ({HRJOB.JB_STATUS} = 'INAC' or Left({HRJOB.JB_DESCR}, 2) = 'Z ')"
    End If
End If

If Len(glbstrSelCri) > 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
'End If
'Modified by Franks on Feb 12,2002

'If Len(glbstrSelCri) >= 0 Then
'    Me.vbxCrystal.SelectionFormula = glbstrSelCri
'End If


' window title if appropriate
Me.vbxCrystal.WindowTitle = "Position Master Report"

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

Private Sub Cri_Skills()
'Laura
Dim EECri As String, OneSet%, X%
Dim strC2, strCx As String
Dim strCa$

OneSet% = False

For X% = 4 To 4
    If Len(clpcode(X%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next X%

If OneSet% = 0 Then Exit Sub

strCa$ = "HRJOBSKL.JS_SKILL"

For X% = 4 To 4
  If Len(clpcode(X%).Text) > 0 Then
    EECri = "({" & strCa$ & "} in  ['" & Replace(clpcode(X%).Text, ",", "','") & "'])"
  End If
Next X%

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
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
' first set primary grouping
GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{HRJOB.JB_GRPCD}" 'changed by RAUBREY 5/9/97
    'Case 0: grpField$ = "{HRJOB.JB_DESCR}"
    Case 1: grpField$ = "{HRJOB.JB_STATUS}"
    Case 2: grpField$ = "{HRJOB.JB_ORG}"
    Case Else:
        If optGrouping(3) Then
            If GrpIdx% = 3 Then
                grpField$ = "{HRJOBEVL.JE_LEVEL}"
            Else
                grpField$ = "{HRJOB.JB_COMPNO}"
            End If
        Else
            grpField$ = "{HRJOB.JB_COMPNO}"
        End If
End Select

dscGroup$ = comGroup(1).Text
dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'" '"descGroup" & CStr(2) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(1) = dscGroup$
If optGrouping(3) Then
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblPosGroup.TB_DESC}" 'changed by RAUBREY 5/9/97
        Case 1: grpField$ = "{tblPosStatus.TB_DESC}"
        Case 2: grpField$ = "{HRTABL.TB_DESC}"
        Case 3: grpField$ = "{HRJOBEVL.JE_LEVEL}"
        Case Else
                grpField$ = "{HRJOB.JB_COMPNO}"
    End Select
End If
grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
'grpCond$ = "GROUP" & CStr(1) & ";{};ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
If grpField$ = "{HRJOB.JB_COMPNO}" Then 'GrpIdx% = 3 Then
    strSFormat$ = "GF1;F;F;F;F;F;F;F"
Else
    strSFormat$ = "GF1;T;X;F;X;X;X;X"
End If
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
'Grouping 2
z% = z% + 1
GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{HRJOB.JB_ORG}"
    Case 1: grpField$ = "{HRJOB.JB_REPTAU}"
    Case 2: grpField$ = "{HRJOB.JB_GRPCD}"
    Case Else:
    'grpField$ = "{HRJOB.JB_DESCR}"
        If optGrouping(3) Then
            If GrpIdx% = 3 Then
                grpField$ = "{HRJOBEVL.JE_LEVEL}"
            Else
                grpField$ = "{HRJOB.JB_DESCR}"
            End If
        Else
            grpField$ = "{HRJOB.JB_DESCR}"
        End If
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$
'
' FINAL SORT
GrpIdx% = comGroup(3).ListIndex
GrpIdx% = 0  'Final Sort = Zero (By Description)
Select Case GrpIdx%
    'Case 0: grpField$ = "{HRJOB.JB_REPTAU}"
    Case 0: grpField$ = "{HRJOB.JB_DESCR}"
End Select
grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(2) = grpCond$

Cri_Sorts = z% ' next section number to format
If optGrouping(4).Value = False Then
    Call setRptLabel(Me, 0)
    Dim xPosGroup
    If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221
        xPosGroup = "Salary Level"
    Else
        xPosGroup = lStr("Position Group")
    End If
    If Left(xPosGroup, 9) = "Position " Then xPosGroup = Mid(xPosGroup, 10)
    vbxCrystal.Formulas(101) = "lblPosGroup='" & xPosGroup & "'"
    Dim xPosStatus
    If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #15590
        xPosStatus = "Income Code"
    Else
        xPosStatus = lStr("Position Status")
    End If
    If Left(xPosStatus, 9) = "Position " Then xPosStatus = Mid(xPosStatus, 10)
    vbxCrystal.Formulas(102) = "lblPosStatus='" & xPosStatus & "'"
    If glbMultiGrid Then
        vbxCrystal.Formulas(103) = "lblGridCategory='" & lStr("Grid Category") & "'"
    End If
End If
End Function

Private Function CriCheck()
Dim X%

CriCheck = False
For X% = 0 To 1

    If Len(clpJob(X%).Text) > 0 And clpJob(X%).Caption = "Unassigned" Then
        MsgBox "If Job is entered - it must be known"
         clpJob(X%).SetFocus
        Exit Function
    End If
Next X%


For X% = 1 To 5 '11
    If Not clpcode(X).ListChecker Then Exit Function

Next X%

If Len(clpNationalClass) > 0 And clpNationalClass.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpNationalClass.SetFocus
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
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

'Ticket #16228 - Oxford
If glbCompSerial = "S/N - 2259W" Then
    glbMulti = True
End If

Call comGrpLoad
If glbSyndesis Then
    lblBCode(2).Caption = "Position Grade"
    clpcode(1).Tag = "00-Enter Position Grade"
End If
If glbCompSerial = "S/N - 2172W" Then 'Lanark
    lblBCode(2).Caption = "Salary Level"
    clpcode(1).Tag = "00-Enter Salary Level"
    clpcode(1).TABLTitle = "Salary Level Code"
End If
If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If

If glbCompSerial = "S/N - 2279W" Then   'Friesens Corporation - Ticket #16189
    'Show Machine Code fields
    clpcode(0).Visible = True
    lblMachine.Visible = True
End If

Call ShowHide 'laura dec 05, 1997
optGrouping(4).Visible = glbCompSerial = "S/N - 2191W"

Call setRptCaption(Me)
Call INI_Controls(Me)

If glbCompSerial = "S/N - 2259W" Then   'Oxford Ticket #15590
    lblStatus.Caption = "Income Code"
    clpcode(3).TABLTitle = "Income Code"
End If
If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob(0).TransDiv = glbWFCUserSecList
    clpJob(1).TransDiv = glbWFCUserSecList
    clpJob(0).TextBoxWidth = 1265 'Ticket #25911 Franks 01/27/2015
    clpJob(1).TextBoxWidth = 1265 'Ticket #25911 Franks 01/27/2015
End If
'Ticket #26726 Franks 06/17/2015 - open it for all
clpJob(0).TextBoxWidth = 1265

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub



Private Sub optGrouping_Click(Index As Integer, Value As Integer)
Call ShowHide   'laura dec 05, 1997
frePaper.Visible = optGrouping(4)

If glbWFC Then 'Ticket #27531 Franks 09/15/2015
    Call WFCScreenSetup
End If

End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call ShowHide   'laura dec 05, 1997

Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optGrouping_LostFocus(Index As Integer)
Call ShowHide   'laura dec 05, 1997

End Sub


Private Sub ShowHide()
Dim X As Integer

If optGrouping(0) Or optGrouping(1) Or optGrouping(4) Then
  For X = 4 To 5 '
      clpcode(X).Text = ""
      clpcode(X).Visible = False
  Next X
  lblTitle(0).Visible = False
  lblEval.Visible = False
  lblLevel.Visible = False
  medLevel.Visible = False

ElseIf optGrouping(2).Value = True Then
  For X = 5 To 5 ' 11
      clpcode(X).Text = ""
      clpcode(X).Visible = False
  Next X
  lblEval.Visible = False
  lblLevel.Visible = False
  medLevel.Visible = False
  
  For X = 4 To 4
      clpcode(X) = ""
      clpcode(X).Visible = True
  Next X
  lblTitle(0).Visible = True

ElseIf optGrouping(3).Value = True Then
  For X = 4 To 4
      clpcode(X).Text = ""
      clpcode(X).Visible = False
  Next X
  lblTitle(0).Visible = False

  For X = 5 To 5 '11
      clpcode(X) = ""
      clpcode(X).Visible = True
  Next X
  lblEval.Visible = True
  
  lblLevel.Top = lblGrid.Top
  medLevel.Top = clpGrid.Top
  lblLevel.Visible = True
  medLevel.Visible = True
  lblEval.Top = lblTitle(0).Top
  clpcode(5).Top = clpcode(4).Top
End If

Call comGrpLoad

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

Function mod_Upd_Pos_Totals(updPCtComp%)
Dim snapJobCount As New ADODB.Recordset
Dim rsHRJob As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#
Dim snapFTENum As New ADODB.Recordset
Dim FTEHrs&
Dim snapFTEHrs As New ADODB.Recordset

Dim spct%

mod_Upd_Pos_Totals = False
On Error GoTo mod_Upd_Pos_Totals_Err
MDIMain.panHelp(0).FloodShowPct = True
'MDIMain.panHelp(0).ForeColor = &HFFFFFF
pct# = 1
If updPCtComp Then
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = pct#
End If

pct# = 3
SQLQ = "UPDATE HRJOB "
SQLQ = SQLQ & " SET HRJOB.JB_NBRFIL = 0, "
SQLQ = SQLQ & " HRJOB.JB_POINTS = 0, "
SQLQ = SQLQ & " HRJOB.JB_FTETOTNU = 0, "
SQLQ = SQLQ & " HRJOB.JB_FTETOTHR = 0 "

'gdbIhr001.Execute (SQLQ$) ' zero out existing numbers
gdbAdoIhr001.Execute SQLQ
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If

'--problem: some employee's records in HR_JOB_HISTORY have no JH_COMPNO, some have
'--so one position has two records in qry_Count_Pos_Filled, the Positions Filled maybe wrong
'--we don't group by JH_COMPNO, JH_JOB, only by JH_JOB
'"qry_Count_Pos_Filled"
''SELECT     JH_COMPNO, JH_JOB, COUNT(JH_EMPNBR) AS NoPosFilled
''From dbo.HR_JOB_HISTORY
''Where (JH_CURRENT <> 0)
''GROUP BY JH_COMPNO, JH_JOB
'SQLQ = "qry_Count_Pos_Filled"
SQLQ = "SELECT JH_JOB, COUNT(JH_EMPNBR) AS NoPosFilled FROM HR_JOB_HISTORY WHERE (JH_CURRENT <> 0) GROUP BY JH_JOB"

snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 5
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If


rsHRJob.Open "HRJOB", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'rsHrJob.Index = "PrimaryKey"
If Not snapJobCount.BOF And Not snapJobCount.EOF Then
    snapJobCount.MoveLast
    rcount& = snapJobCount.RecordCount
    snapJobCount.MoveFirst
    ipct# = 20 / rcount&
End If



While Not snapJobCount.BOF And Not snapJobCount.EOF
    Job$ = snapJobCount("JH_JOB")
    JobCount& = snapJobCount("NoPosFilled")
    rsHRJob.MoveFirst
    rsHRJob.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJob.EOF Then
        'rsHrJob.Edit
        rsHRJob("JB_NBRFIL") = JobCount&
        rsHRJob.Update
    End If

    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
        
    End If

    snapJobCount.MoveNext
Wend

snapJobCount.Close


SQLQ = "qry_Count_Pos_Points"

If snapEvalPoints.State <> 0 Then snapEvalPoints.Close
snapEvalPoints.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 25
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If

If Not snapEvalPoints.BOF And Not snapEvalPoints.EOF Then
    snapEvalPoints.MoveLast
    rcount& = snapEvalPoints.RecordCount
    snapEvalPoints.MoveFirst
    ipct# = 20 / rcount&
End If



While Not snapEvalPoints.BOF And Not snapEvalPoints.EOF
    Job$ = snapEvalPoints("JE_CODE")
    JobPoints# = snapEvalPoints("TotalPoints")
    
    'rsHrJob.Seek "=", Job$
    rsHRJob.MoveFirst
    rsHRJob.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJob.EOF Then
        'rsHrJob.Edit
        rsHRJob("JB_POINTS") = JobPoints#
        rsHRJob.Update
    End If

    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
    End If

    snapEvalPoints.MoveNext
Wend

snapEvalPoints.Close

SQLQ = "qry_Count_FTENum"

snapFTENum.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 45
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If

If Not snapFTENum.BOF And Not snapFTENum.EOF Then
    snapFTENum.MoveLast
    rcount& = snapFTENum.RecordCount
    snapFTENum.MoveFirst
    ipct# = 20 / rcount&
End If



While Not snapFTENum.BOF And Not snapFTENum.EOF
    Job$ = snapFTENum("JH_JOB")
    If IsNull(snapFTENum("FTENumTot")) Then    'laura 03/05/98
      FTENum# = 0
    Else
      FTENum# = snapFTENum("FTENumTot")
    End If
    
    'rsHrJob.Seek "=", Job$
    rsHRJob.MoveFirst
    rsHRJob.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJob.EOF Then
        'rsHrJob.Edit
        rsHRJob("JB_FTETotNu") = FTENum#
        rsHRJob.Update
    End If

    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
    End If

    snapFTENum.MoveNext
Wend

snapFTENum.Close

SQLQ = "qry_Count_FTEHrs"

snapFTEHrs.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 65
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If

If Not snapFTEHrs.BOF And Not snapFTEHrs.EOF Then
    snapFTEHrs.MoveLast
    rcount& = snapFTEHrs.RecordCount
    snapFTEHrs.MoveFirst
    ipct# = 30 / rcount&
End If



While Not snapFTEHrs.BOF And Not snapFTEHrs.EOF
    Job$ = snapFTEHrs("JH_JOB")
    If IsNull(snapFTEHrs("FTEHrsTot")) Then     'laura 03/04/98
      FTEHrs& = 0
    Else
      FTEHrs& = snapFTEHrs("FTEHrsTot")
    End If
    
    rsHRJob.MoveFirst
    rsHRJob.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJob.EOF Then
        rsHRJob("JB_FTETotHr") = FTEHrs&
        rsHRJob.Update
    End If

    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
    End If

    snapFTEHrs.MoveNext
Wend
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = 99
End If

snapFTEHrs.Close

rsHRJob.Close


MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(0).ForeColor = &H0&
MDIMain.panHelp(0).FloodType = 0
mod_Upd_Pos_Totals = True

Exit Function


mod_Upd_Pos_Totals_Err:
If Err = 94 Then
Err = 0
Resume Next
End If
glbFrmCaption$ = "Module - Count Positions"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update HRJOB/Eval Count", "HRJOB/eval/skls", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function


'Private Sub Cri_Machine()
'Dim EECri As String
'
'If Len(clpcode(0).Text) > 0 Then
'    EECri = "{HRJOB.JB_MACHINE_NUM} = '" & clpcode(0).Text & "' "
'    HisSQL = HisSQL & " AND JB_MACHINE_NUM='" & clpcode(0).Text & "' "
'End If
'
'
'If Len(EECri) >= 1 Then
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'    glbiOneWhere = True
'End If
'End Sub
'

Private Sub WFCScreenSetup() 'Ticket #27531 Franks 09/14/2015
    If optGrouping(1).Value Then
        lblJob.Visible = True
        clpJobMaster.Text = ""
        clpJobMaster.Visible = True
    Else
        lblJob.Visible = False
        clpJobMaster.Visible = False
    End If
End Sub

Private Sub Cri_JobMaster() 'Ticket #27531 Franks 09/15/2015
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String

If Len(clpJobMaster.Text) <= 0 Then Exit Sub

PosCri = "({HRJOB.JB_JOBCODE} = '" & clpJobMaster.Text & "')"

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If
End Sub

