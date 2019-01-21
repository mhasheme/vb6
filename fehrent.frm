VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmHrEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Hourly Entitlements"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   1470
   ClientWidth     =   11880
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
   ScaleHeight     =   7725
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame frDays 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   2840
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
      Begin MSMask.MaskEdBox medEntitleAmntD 
         DataField       =   "HE_ENTITLEDAY"
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Tag             =   "21-Amount of entitlement during the period"
         Top             =   360
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
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPrevAmntD 
         DataField       =   "HE_PREVDAY"
         Height          =   285
         Left            =   0
         TabIndex        =   37
         Tag             =   "21-Amount of Previous Year Outstanding entitlement"
         Top             =   0
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
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTakenAmntD 
         DataField       =   "HE_TAKENDAY"
         Height          =   285
         Left            =   0
         TabIndex        =   38
         Tag             =   "21-Amount of Previous Year Outstanding entitlement"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "##,##0.00;(##,##0.00)"
         PromptChar      =   "_"
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehrent.frx":0000
      Height          =   2595
      Left            =   120
      OleObjectBlob   =   "fehrent.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   10335
   End
   Begin INFOHR_Controls.DateLookup dlpTDate 
      DataField       =   "HE_TDATE"
      Height          =   285
      Left            =   6540
      TabIndex        =   3
      Tag             =   "41-Ending date"
      Top             =   3960
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFDate 
      DataField       =   "HE_FDATE"
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Tag             =   "41-Starting date"
      Top             =   3960
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "HE_TYPE"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Tag             =   "01-Entitlement - Code"
      Top             =   3600
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado2"
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   7065
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
      Begin VB.CommandButton cmdDays 
         Appearance      =   0  'Flat
         Caption         =   "Da&ys"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Tag             =   "Display Hourly Entitlements in Days"
         Top             =   0
         Width           =   875
      End
      Begin VB.CommandButton cmdHours 
         Appearance      =   0  'Flat
         Caption         =   "&Hours"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   29
         Tag             =   "Display Hourly Entitlements in Hours"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdFlag 
         Caption         =   "Flag &Updt"
         Height          =   375
         Left            =   4320
         TabIndex        =   25
         Tag             =   "Mass Update of COE Flags"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdReCalc 
         Caption         =   "&ReCalculate"
         Height          =   375
         Left            =   2640
         TabIndex        =   24
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "HE_LDATE"
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
      Index           =   0
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "HE_LTIME"
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
      Index           =   1
      Left            =   4320
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "HE_LUSER"
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
      Index           =   2
      Left            =   6000
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
         Left            =   7440
         TabIndex        =   26
         Top             =   137
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   137
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Employee Name"
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
         Left            =   3120
         TabIndex        =   12
         Top             =   137
         Width           =   1740
      End
   End
   Begin Threed.SSCheck chkCOEFlag 
      DataField       =   "HE_COE"
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Tag             =   "Include entitlement on COE report ?"
      Top             =   5400
      Width           =   2925
      _Version        =   65536
      _ExtentX        =   5159
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Cost of Employment                          "
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
   Begin MSMask.MaskEdBox medEntitleAmnt 
      DataField       =   "HE_ENTITLE"
      Height          =   285
      Left            =   2835
      TabIndex        =   5
      Tag             =   "21-Amount of entitlement during the period"
      Top             =   4680
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
      Format          =   "##,##0.00;(##,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDHRS 
      DataField       =   "HE_DHRS"
      Height          =   285
      Left            =   7920
      TabIndex        =   21
      Tag             =   "Hours per Day"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
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
      Format          =   "###.00;(###.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtTaken 
      DataField       =   "He_Taken"
      Height          =   285
      Left            =   7920
      TabIndex        =   22
      Tag             =   "21-Amount of entitlement during the period"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
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
      Format          =   "##,###.00;(##,###.00)"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8040
      Top             =   6840
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
   Begin MSMask.MaskEdBox medPrevAmnt 
      DataField       =   "HE_PREV"
      Height          =   285
      Left            =   2835
      TabIndex        =   4
      Tag             =   "21-Amount of Previous Year Outstanding entitlement"
      Top             =   4320
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
      Format          =   "##,##0.00;(##,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "HE_TAKEN"
      Height          =   285
      Left            =   2835
      TabIndex        =   6
      Tag             =   "21-Amount of Previous Year Outstanding entitlement"
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
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
      Format          =   "##,##0.00;(##,##0.00)"
      PromptChar      =   "_"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
      Bindings        =   "fehrent.frx":4EF0
      Height          =   2595
      Left            =   120
      OleObjectBlob   =   "fehrent.frx":4F04
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.Label lblHrsDaysTaken 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
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
      Left            =   4200
      TabIndex        =   33
      Top             =   5085
      Width           =   420
   End
   Begin VB.Label lblHrsDaysEntit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
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
      Left            =   4200
      TabIndex        =   32
      Top             =   4725
      Width           =   420
   End
   Begin VB.Label lblHrsDaysPrv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
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
      Left            =   4200
      TabIndex        =   31
      Top             =   4365
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Taken"
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
      Left            =   240
      TabIndex        =   28
      Top             =   5085
      Width           =   465
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
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
      Left            =   240
      TabIndex        =   27
      Top             =   4365
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   5640
      TabIndex        =   20
      Top             =   4005
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   4725
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   4005
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3645
      Width           =   960
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "HE_EMPNBR"
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
      Height          =   195
      Left            =   1470
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "HE_COMPNO"
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
      Height          =   195
      Left            =   660
      TabIndex        =   16
      Top             =   6390
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmHrEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew%
Dim fUPMode As Integer ', fglbEmptyNew As Integer
Dim savDHrs, frmwdate, frmFDate, frmTDate
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim SavFdate, SavTdate
Dim oldEntitleAmnt, oPHrl

Private Sub chkCOEFlag_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkEODollar()
Dim SQLQ As String, Msg As String, dd#

chkEODollar = False

On Error GoTo chkEODollar_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Entitlement code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Entitlement code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

'If FLEX Logic - Cannot create Entitlement logic with '-' suffixed Entitlement Code
If Len(clpCode(1).Text) > 0 Then
    If Right(clpCode(1).Text, 1) = "-" Then
        MsgBox "Invalid Entitlement code. The Entitlement Code cannot have '-' suffixed to it."
        clpCode(1).SetFocus
        Exit Function
    End If
End If

'7.9 Enhancement - Cannot create OTs or CTs code entitlements
If Len(clpCode(1).Text) > 0 Then
    If Left(clpCode(1).Text, 2) = "OT" Or Left(clpCode(1).Text, 2) = "CT" Then
        MsgBox "Invalid Entitlement code. The Hourly Entitlement cannot be set for codes with 'OT' or 'CT' prefix to it."
        clpCode(1).SetFocus
        Exit Function
    End If
End If

'So FINALLY found out the actual reason for preventing users to create VAC, SIC and OT/CT prefixed
'hourly entitlement. This is purely because of how the code has been written in ESS/TS. This more of the
'design issue and can easily be fixed but it will be time consuming. The logic to compute and display the
'outstanding entitlement is built on IF...THEN...ELSE logic. The system first checks for VAC prefixed
'codes and then in the ELSE, SIC prefixed and then in the ELSE OT/CT prefixed and lastely in the ELSE
'hourly entitlement codes. So on the Hourly Entitlement page if VAC prefixed hourly entitle code is found
'then it will first go in the IF VAC statement and check/compute outstanding as in HREMP and not in
'HRENTHRS table hence giving wrong value.

'This is just a note as I keep on forgetting why this was added whenever a question is asked. This was added
'with the request from Jerry after Mostafa explained to Jerry when VAC or SIC codes are used for Hourly
'Entitlement, ESS and TS showing the outstanding balance for Hourly Entitlement gets messed up with VAC
'and SIC built-in logic in ESS/TS. Mostafa also confirmed that he himself does not now remember the other
'reasons connected to this.
'Please note: VAC or SIC prefixed Hourly Entitlement codes in info:HR does not create any issues. This is
'purely for ESS and TS.
'7.9 Enhancement - Warn to not create VACs or SICs code entitlements. This is for those client using ESS/TS.
If Len(clpCode(1).Text) > 0 Then
    If Left(clpCode(1).Text, 3) = "VAC" Or Left(clpCode(1).Text, 3) = "SIC" Then
        MsgBox "Please avoid creating Hourly Entitlements for codes prefixed 'VAC' and 'SIC', if using ESS/Timesheet Web Modules.", vbExclamation, "info:HR - Hourly Entitlement"
        'clpCode(1).SetFocus
        'Exit Function
    End If
End If

If Len(dlpFDate.Text) >= 1 Then
    If Not IsDate(dlpFDate.Text) Then
        MsgBox "From Date is not a valid date."
        dlpFDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "From Date is required."
    dlpFDate.SetFocus
    Exit Function
End If

If Len(dlpTDate.Text) >= 1 Then
    If Not IsDate(dlpTDate.Text) Then
        MsgBox "To Date is not a valid date."
        dlpTDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "To Date is required."
    dlpTDate.SetFocus
    Exit Function
End If

dd# = DateDiff("d", CVDate(dlpFDate.Text), CVDate(dlpTDate.Text))
If dd# < 0 Then
    MsgBox "From date must be earlier than To Date"
    dlpFDate.SetFocus
    Exit Function
End If

If Len(medPrevAmnt) = 0 Then medPrevAmnt = 0
If Not IsNumeric(medPrevAmnt) Then
    MsgBox "Invalid Previous year outstanding entitlement"
    medPrevAmnt.SetFocus
    Exit Function
End If

'If FLEX Logic then Entitlement Amount is not entered
If Len(clpCode(1).Text) > 0 Then
    If Right(clpCode(1).Text, 1) <> "+" Or fglbNew% Then
        If Len(Trim(medEntitleAmnt)) = 0 Then
            MsgBox "Entitlement Amount is required."
            medEntitleAmnt.SetFocus
            Exit Function
        End If
        If Not IsNumeric(medEntitleAmnt) Then
            MsgBox "Invalid Entitlement Amount"
            medEntitleAmnt.SetFocus
            Exit Function
        End If
    End If
End If

If Not ChkDup Then
    MsgBox "Duplicate record existed - not entered"
    clpCode(1).SetFocus
    Exit Function
End If
chkEODollar = True

Exit Function

chkEODollar_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkODollar", "HRENTHRS", "edit/Add")
Resume Next

End Function

Private Function ChkDup()
Dim rsHE As New ADODB.Recordset
Dim SQLQ
SQLQ = ""
ChkDup = False
If glbtermopen Then
    SQLQ = "SELECT HE_ID FROM Term_ENTHRS WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " AND HE_TYPE='" & clpCode(1) & "'"
    SQLQ = SQLQ & " AND HE_TDATE=" & Date_SQL(dlpTDate)
    If Not fglbNew Then
        SQLQ = SQLQ & " AND HE_ID <> " & Data1.Recordset("HE_ID")
    End If
    rsHE.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
Else
    SQLQ = "SELECT HE_ID FROM HRENTHRS WHERE HE_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " AND HE_TYPE='" & clpCode(1) & "'"
    SQLQ = SQLQ & " AND HE_TDATE=" & Date_SQL(dlpTDate)
    If Not fglbNew Then
        SQLQ = SQLQ & " AND HE_ID <> " & Data1.Recordset("HE_ID")
    End If
    rsHE.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
End If
If Not rsHE.EOF Then Exit Function
ChkDup = True
End Function

Sub cmdCancel_Click()
Dim SQLQ, x

On Error GoTo Can_Err

rsDATA.CancelUpdate

Call Display_Value


fglbNew% = False
'Call ST_UPD_MODE(True)  ' reset screen's attributes
Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRENTHRS", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMHRENT" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x
Dim xID, xKey
Dim rzAttend As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim SQLQ As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

Call Append_Accrual(glbLEE_ID, clpCode(1), Date, 0 - (medEntitleAmnt + medPrevAmnt), "X", "Deleted the Hourly Entitlement")

'I am taking this delete from Attendance off because there will be other attendance records for this + code
'that is also being used to compute the Entitlement amount. If I delete only the opening balance record then
'what will happen to the rest of the + code records. I cannot delete it as the user may want to keep those
'records but just delete the hourly entitlement. Besides the Accrual table is already being taken care about.
'The user will manually delete it if required.
''Ticket #21067 - If the Entitlement Code is suffixed with + then delete the corresponding Attendance record
''for the Hourly Entitlement earned
'If Right(clpCode(1).Text, 1) = "+" Then
'    'Delete Record in Attendance screen
'
'    'Retrieve the hire date from employee table
'    If Not glbtermopen Then
'        SQLQ = "SELECT ED_EMPNBR,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
'        rsHrEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'    Else
'        SQLQ = "SELECT ED_EMPNBR,ED_DOH FROM TERM_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
'        rsHrEmp.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
'    End If
'
'    If Not glbtermopen Then
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & glbLEE_ID
'    Else
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & glbTERM_ID & " AND TERM_SEQ =" & glbTERM_Seq
'    End If
'    SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(1).Text & "'"
'    If CVDate(rsHrEmp("ED_DOH")) > CVDate(dlpFDate.Text) Then
'        SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(rsHrEmp("ED_DOH"))
'    Else
'        SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFDate.Text)
'    End If
'    'SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFDate.Text)
'    If Not glbtermopen Then
'        rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Else
'        rzAttend.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'    End If
'    If Not rzAttend.EOF Then
'        rzAttend.Delete
'    End If
'    rsHrEmp.Close
'    Set rsHrEmp = Nothing
'    rzAttend.Close
'    Set rzAttend = Nothing
'End If
    
xID = Data1.Recordset("HE_ID")

xKey = Data1.Recordset("HE_EMPNBR")
xKey = xKey & "|" & Format(Data1.Recordset("HE_FDATE"), "dd-mmm-yyyy")
xKey = xKey & "|" & Format(Data1.Recordset("HE_TDATE"), "dd-mmm-yyyy")
xKey = xKey & "|" & Data1.Recordset("HE_TYPE")
xKey = xKey & "|"
xKey = xKey & "|" & Format(Data1.Recordset("HE_FDATE"), "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
Call Entitlements_Master_Integration(xKey, , True)
If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRENTHRS", "Delete")

Resume Next

Unload Me

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If gSec_Upd_Hrly_Entitlements Then
        If Index = 1 Then
            'FLEX Logic only - Disable the Entitlement Amount change
            If Len(clpCode(1).Text) > 0 Then
                If Right(clpCode(1).Text, 1) = "+" Then
                    'Ticket #21067 - Only allow if adding a new record. This will also create an Attendance
                    'record of the entitlement amount. Any modification to the entitlement amount should
                    'be done from Attendance screen thereafter.
                    If fglbNew% Then
                        medEntitleAmnt.Enabled = True
                    Else
                        medEntitleAmnt.Enabled = False
                    End If
                Else
                    medEntitleAmnt.Enabled = True
                End If
            Else
                medEntitleAmnt.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmdDays_Click()
    cmdDays.Enabled = False
    cmdHours.Enabled = True
    lblHrsDaysEntit.Caption = "Days"
    lblHrsDaysPrv.Caption = "Days"
    lblHrsDaysTaken.Caption = "Days"
    frDays.Visible = True
    vbxTrueGrid1.Visible = True
    vbxTrueGrid.Visible = False
End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdFlag_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant
Dim rsHR As New ADODB.Recordset
On Error GoTo MAll_Err

Msg$ = "How would you like to mark all COE flags?"
Msg = Msg$
Title$ = "Mark all completed?"   ' zzz
DgDef = MB_YESNOCANCEL + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If glbtermopen Then
    rsHR.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    rsHR.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !HE_COE = True
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    Data1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT
End If

If Response = IDNO Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !HE_COE = False
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    Data1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT
End If

Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HRENTHRS", "Mark All")
Call RollBack '28July99 js

End Sub

Private Sub cmdFlag_GotFocus()
Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

If Not IsNumeric(savDHrs) Then savDHrs = 0
txtDHRS = savDHrs
oldEntitleAmnt = medEntitleAmnt
SavFdate = dlpFDate
SavTdate = dlpTDate
oPHrl = medPrevAmnt

'FLEX Logic only - Disable the Entitlement Amount change
If gSec_Upd_Hrly_Entitlements Then
    If Len(clpCode(1).Text) > 0 Then
        If Right(clpCode(1).Text, 1) = "+" Then
            medEntitleAmnt.Enabled = False
        Else
            medEntitleAmnt.Enabled = True
        End If
    Else
        medEntitleAmnt.Enabled = True
    End If
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRENTHRS", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err

fglbNew% = True

Call Set_Control("B", Me)
rsDATA.AddNew

If Not IsNumeric(savDHrs) Then savDHrs = 0
txtDHRS = savDHrs
dlpFDate.Text = ""
dlpTDate.Text = ""
dlpFDate.Text = frmFDate
dlpTDate.Text = frmTDate
chkCOEFlag.Value = True
SavFdate = ""
SavTdate = ""
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"
TxtTaken.Text = 0 'Ticket #10918
clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRENTHRS", "Add")
Resume Next

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x, SQLQ
Dim xID
Dim xKey
Dim xComments As String
Dim xDiffHours

On Error GoTo Add_Err

If Not chkEODollar() Then Exit Sub

If fglbNew Then
    'xKey = 0
    xKey = glbLEE_ID
    xKey = xKey & "|" & Format(dlpFDate, "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpTDate, "dd-mmm-yyyy")
    'xKey = xKey & "|ANY"
    xKey = xKey & "|" & clpCode(1).Text
    xKey = xKey & "|" & medEntitleAmnt
Else
    xKey = Data1.Recordset("HE_EMPNBR")
    xKey = xKey & "|" & Format(Data1.Recordset("HE_FDATE"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(Data1.Recordset("HE_TDATE"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Data1.Recordset("HE_TYPE")
End If

If Not glbtermopen Then
    Call UpdUStats(Me)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA!HE_ID
    
    'Ticket #21067 - For New Record, if the Entitlement Code is suffixed with + then insert an Attendance
    'record for the Hourly Entitlement earned - helps in the Recalculate function
    If fglbNew Then
        Call Add_Attendance_for_FlexCode
    End If
    
    Call EmpReCalc
    
    'Ticket #17924 - Incase the date range is changed manually - the entitlement should be recalculated.
    If dlpTDate.Text <> SavTdate Or dlpFDate.Text <> SavFdate Then
        Screen.MousePointer = HOURGLASS
        Call EntReCalcHr(glbLEE_ID)
        Screen.MousePointer = DEFAULT
    End If
Else
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    xID = rsDATA!HE_ID
    gdbAdoIhr001X.CommitTrans
    
    'Ticket #21067 - For New Record, if the Entitlement Code is suffixed with + then insert an Attendance
    'record for the Hourly Entitlement earned - helps in the Recalculate function
    If fglbNew Then
        Call Add_Attendance_for_FlexCode
    End If
    
    If glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Ticket #23417 Franks 06/03/2013
        Call EntReCalcHrTerm(glbTERM_Seq)
    End If
End If

Data1.Refresh

If xID > 0 Then
    Data1.Recordset.Find "HE_ID=" & xID
End If

If Not glbtermopen Then
    If fglbNew% Then
        Call Append_Accrual(glbLEE_ID, clpCode(1), dlpFDate, medEntitleAmnt, "H", "Added Hourly Entitlement")
    Else
        'Previous
        'Ticket #23001 - Adding appropriate update for Prev. Hourly Entitlement to Accrual file
        xDiffHours = Val(medPrevAmnt) - Val(oPHrl)
        If xDiffHours <> 0 Then
            xComments = "Prev. Hourly Ent. Chg from " & oPHrl & " to " & medPrevAmnt & ". OS: " & (Val(IIf(IsNull(oPHrl), 0, oPHrl)) + Val(IIf(IsNull(oldEntitleAmnt), 0, oldEntitleAmnt))) - Val(IIf(IsNull(MaskEdBox1), 0, MaskEdBox1))
            If CVDate(Date) >= CVDate(dlpFDate.Text) And CVDate(Date) <= CVDate(dlpTDate.Text) Then
                Call Append_Accrual(glbLEE_ID, clpCode(1), Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(glbLEE_ID, clpCode(1), dlpTDate.Text, xDiffHours, "C", xComments)
            End If
        End If
    
        'Current
        xDiffHours = Val(medEntitleAmnt) - Val(oldEntitleAmnt)
        If xDiffHours <> 0 Then
            xComments = "Current. Hourly Ent. Chg from " & oldEntitleAmnt & " to " & medEntitleAmnt & ". OS: " & (Val(IIf(IsNull(oPHrl), 0, oPHrl)) + Val(IIf(IsNull(oldEntitleAmnt), 0, oldEntitleAmnt))) - Val(IIf(IsNull(MaskEdBox1), 0, MaskEdBox1))
            If CVDate(Date) >= CVDate(dlpFDate.Text) And CVDate(Date) <= CVDate(dlpTDate.Text) Then
                'Call Append_Accrual(glbLEE_ID, clpCode(1), Date, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", "Changed the current Hourly Entitlement")
                Call Append_Accrual(glbLEE_ID, clpCode(1), Date, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", xComments)
            Else
                'Call Append_Accrual(glbLEE_ID, clpCode(1), dlpTDate.Text, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", "Changed the current Hourly Entitlement")
                Call Append_Accrual(glbLEE_ID, clpCode(1), dlpTDate.Text, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", xComments)
            End If
        End If
    End If
Else
    If fglbNew% Then
        Call Append_Accrual(glbTERM_ID, clpCode(1), dlpFDate, medEntitleAmnt, "H", "Added Hourly Entitlement")
    Else
        'Previous
        'Ticket #23001 - Adding appropriate update for Prev. Hourly Entitlement to Accrual file
        xDiffHours = Val(medPrevAmnt) - Val(oPHrl)
        If xDiffHours <> 0 Then
            xComments = "Prev. Hourly Ent. Chg from " & oPHrl & " to " & medPrevAmnt & ". OS: " & (Val(IIf(IsNull(oPHrl), 0, oPHrl)) + Val(IIf(IsNull(oldEntitleAmnt), 0, oldEntitleAmnt))) - Val(IIf(IsNull(MaskEdBox1), 0, MaskEdBox1))
            If CVDate(Date) >= CVDate(dlpFDate.Text) And CVDate(Date) <= CVDate(dlpTDate.Text) Then
                Call Append_Accrual(glbLEE_ID, clpCode(1), Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(glbTERM_ID, clpCode(1), dlpTDate.Text, xDiffHours, "C", xComments)
            End If
        End If
    
        'Current
        xDiffHours = Val(medEntitleAmnt) - Val(oldEntitleAmnt)
        If xDiffHours <> 0 Then
            xComments = "Current. Hourly Ent. Chg from " & oldEntitleAmnt & " to " & medEntitleAmnt & ". OS: " & (Val(IIf(IsNull(oPHrl), 0, oPHrl)) + Val(IIf(IsNull(oldEntitleAmnt), 0, oldEntitleAmnt))) - Val(IIf(IsNull(MaskEdBox1), 0, MaskEdBox1))
            If CVDate(Date) >= CVDate(dlpFDate.Text) And CVDate(Date) <= CVDate(dlpTDate.Text) Then
                'Call Append_Accrual(glbTERM_ID, clpCode(1), Date, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", "Changed the current Hourly Entitlement")
                Call Append_Accrual(glbTERM_ID, clpCode(1), Date, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", xComments)
            Else
                'Call Append_Accrual(glbTERM_ID, clpCode(1), dlpTDate.Text, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", "Changed the current Hourly Entitlement")
                Call Append_Accrual(glbTERM_ID, clpCode(1), dlpTDate.Text, Val(medEntitleAmnt) - Val(oldEntitleAmnt), "C", xComments)
            End If
        End If
    End If
End If

fglbNew% = False

Call ST_UPD_MODE(True)
Call SET_UP_MODE
Call Entitlements_Master_Integration(xKey, xID)

If NextFormIF("Entitlement") Then
    Call cmdNew_Click
End If
Exit Sub

Add_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRENTHRS", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Hourly Entitlements"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Hourly Entitlements"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub cmdHours_Click()
    cmdDays.Enabled = True
    cmdHours.Enabled = False
    lblHrsDaysEntit.Caption = "Hours"
    lblHrsDaysPrv.Caption = "Hours"
    lblHrsDaysTaken.Caption = "Hours"
    frDays.Visible = False
    vbxTrueGrid1.Visible = False
    vbxTrueGrid.Visible = True

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalc_Click()
Dim Msg, Response, DgDef

Msg = "Do you wish to proceed and recalculate ALL "
Msg = Msg & "Employee's outstanding Hourly entitlements?"
        
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "ReCalculate")
If Response = IDNO Then Exit Sub
If Data1.Recordset.RecordCount > 0 Then
    Screen.MousePointer = HOURGLASS
    Call EntReCalcHr
    Screen.MousePointer = DEFAULT
    Data1.Refresh
End If

glbENTScreen = True

End Sub

Private Sub CmdRecalc_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub

Function EERetrieve()
Dim SQLQ As String, xlen

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If glbtermopen Then
    'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS  from Term_ENTHRS "
    'Ticket #22363 Franks 08/02/2012
    SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "
    
    SQLQ = SQLQ & " FROM Term_ENTHRS "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY HE_EMPNBR, HE_TYPE, HE_FDATE"
Else
    'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS from HRENTHRS"
    'Ticket #22363 Franks 08/02/2012
    SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "
    
    SQLQ = SQLQ & " FROM HRENTHRS"
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY HE_EMPNBR, HE_TYPE, HE_FDATE"
End If

Data1.RecordSource = SQLQ
Data1.Refresh

If glbtermopen Then
    savDHrs = 0
Else
    Dim rsMU As New ADODB.Recordset
    
    SQLQ = "SELECT * FROM qry_MU_Entitle "
    SQLQ = SQLQ & "WHERE ED_EMPNBR = " & glbLEE_ID & " "
    rsMU.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsMU.EOF Then
        savDHrs = 0
    Else
        savDHrs = rsMU("JH_DHRS")
        If Len(savDHrs) = 0 Then savDHrs = 0
    End If
    rsMU.Close
End If

frmFDate = glbCompEdFrom
frmTDate = glbCompEdTo

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EHourRetrieve", "HRENTHRS", "SELECT")
Resume Next
Exit Function

End Function

Private Sub EmpReCalc()
Dim SQLQ, xwk
Dim rsATT As New ADODB.Recordset

If dlpTDate.Text = SavTdate And dlpFDate.Text = SavFdate Then Exit Sub


SQLQ = "SELECT HR_ATTENDANCE.* FROM HR_ATTENDANCE "
SQLQ = SQLQ & "WHERE AD_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(1).Text & "' "
SQLQ = SQLQ & "ORDER BY AD_EMPNBR, AD_REASON"
rsATT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
xwk = 0
Do Until rsATT.EOF
    If rsATT("AD_DOA") >= DateValue(dlpFDate.Text) And rsATT("AD_DOA") <= DateValue(dlpTDate.Text) Then
        xwk = xwk + rsATT("AD_HRS")
    End If
    rsATT.MoveNext
Loop
rsATT.Close
TxtTaken = xwk

End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    Me.cmdModify_Click
    glbOnTop = "FRMHRENT"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMHRENT"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False


glbOnTop = "FRMHRENT"


If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
    cmdReCalc.Visible = False
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

savDHrs = 0
Screen.MousePointer = DEFAULT
If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Ticket #17924 - Only for SQL versions
If Not glbSQL Then
    medPrevAmnt.Enabled = False
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Hourly Entitlements - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value

'Call ST_UPD_MODE(False)             '
If Not gSec_Upd_Hrly_Entitlements Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdFlag.Enabled = False
    cmdReCalc.Enabled = False
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
If frDays.Visible = False Then      'Release 8.0
    Keepfocus = Not isUpdated(Me)
End If
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmHrEnt = Nothing 'carmen may 00
    Call NextForm
End Sub

Private Sub medEntitleAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
cmdReCalc.Enabled = TF
cmdFlag.Enabled = TF

chkCOEFlag.Enabled = TF
medEntitleAmnt.Enabled = TF
clpCode(1).Enabled = TF
dlpFDate.Enabled = TF
dlpTDate.Enabled = TF
medPrevAmnt.Enabled = TF

'vbxTrueGrid.Enabled = FT

If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
End If
End Sub

Private Sub medPrevAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Call SET_UP_MODE
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then
            'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS  from Term_ENTHRS "
            'Ticket #22363 Franks 08/02/2012
            'SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS  FROM Term_ENTHRS "
            'SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
            
            '8.0 - Ticket #22682 - Option to show in Days
            SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS,  "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "

            SQLQ = SQLQ & " FROM Term_ENTHRS "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS  from HRENTHRS"
            'Ticket #22363 Franks 08/02/2012
            'SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS FROM HRENTHRS"
            'SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
            
            '8.0 - Ticket #22682 - Option to show in Days
            SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE HE_TAKEN END)) AS OUTS,  "
            'SQLQ = SQLQ & " ((CASE WHEN HE_PREV IS NULL THEN 0 ELSE HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE HE_TAKEN END)) AS OUTSDAY  "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "
            
            SQLQ = SQLQ & " FROM HRENTHRS"
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
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

Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
SQLQ = Data1.RecordSource
SQLQ = Left(SQLQ, InStr(SQLQ, "ORDER BY") - 1)
SQLQ = SQLQ & " AND HE_ID = " & Data1.Recordset("HE_ID")

    If glbtermopen Then
        'SQLQ = "SELECT * FROM Term_ENTHRS "
        'SQLQ = SQLQ & " WHERE HE_ID = " & Data1.Recordset("HE_ID")
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        'SQLQ = "SELECT * FROM HRENTHRS"
        'SQLQ = SQLQ & " WHERE HE_ID = " & Data1.Recordset!HE_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call SET_UP_MODE
    Call Set_Control("R", Me, rsDATA)
End If

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
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
UpdateRight = gSec_Upd_Hrly_Entitlements
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
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
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmHrEnt.Caption = "Hourly Entitlements - " & Left$(glbLEE_SName, 5)
    frmHrEnt.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub Add_Attendance_for_FlexCode()
    Dim rzAttend As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsCurJobSal As New ADODB.Recordset
    Dim SQLQ As String
    
    'Ticket #21067
    'If the Entitlement Code is suffixed with + then insert an Attendance record
    'for the Hourly Entitlement earned - helps in the Recalculate function
    If Right(clpCode(1).Text, 1) = "+" Then
        'Retrieve the employee data to update the Attendance record with default values
        If Not glbtermopen Then
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO,ED_DOH FROM TERM_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
            rsHREmp.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        
        'Add Record in Attendance screen
        If Not glbtermopen Then
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & glbLEE_ID
        Else
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & glbTERM_ID & " AND TERM_SEQ =" & glbTERM_Seq
        End If
        SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(1).Text & "'"
        'Ticket #18550 - Attendance record date cannot be prior to hire date
        If CVDate(rsHREmp("ED_DOH")) > CVDate(dlpFDate.Text) Then
            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(rsHREmp("ED_DOH"))
        Else
            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFDate.Text)
        End If
        If Not glbtermopen Then
            rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rzAttend.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        End If
        If rzAttend.EOF Then
            rzAttend.AddNew
                
            If glbtermopen Then
                rzAttend("AD_EMPNBR") = glbTERM_ID
                rzAttend("TERM_SEQ") = glbTERM_Seq
            Else
                rzAttend("AD_EMPNBR") = glbLEE_ID
            End If
        End If
        rzAttend("AD_COMPNO") = "001"
        rzAttend("AD_DOA") = dlpFDate.Text
        rzAttend("AD_REASON") = clpCode(1).Text
        rzAttend("AD_HRS") = medEntitleAmnt

        If Not rsHREmp.EOF Then
            rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
            rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
            rzAttend("AD_ORG") = rsHREmp("ED_ORG")
            
            'Ticket #18550 - Attendance record date cannot be prior to hire date
            If CVDate(rsHREmp("ED_DOH")) > CVDate(dlpFDate.Text) Then
                rzAttend("AD_DOA") = rsHREmp("ED_DOH")
            End If
        End If
        rsHREmp.Close

        If Not glbtermopen Then
            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID
            rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM TERM_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbTERM_ID & " AND TERM_SEQ =" & glbTERM_Seq
            rsCurJobSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        If Not rsCurJobSal.BOF Then
            If rsCurJobSal("SH_SALARY") > 0 Then
                rzAttend("AD_SALARY") = rsCurJobSal("SH_SALARY")
                rzAttend("AD_SALCD") = rsCurJobSal("SH_SALCD")
            End If
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        If Not glbtermopen Then
            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbLEE_ID
            rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM TERM_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbTERM_ID & " AND TERM_SEQ =" & glbTERM_Seq
            rsCurJobSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        If Not rsCurJobSal.EOF Then
            rzAttend("AD_JOB") = rsCurJobSal("JH_JOB")
            rzAttend("AD_DHRS") = rsCurJobSal("JH_DHRS")
            rzAttend("AD_WHRS") = rsCurJobSal("JH_WHRS")
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        'Ticket #18550
        'rzAttend("AD_COMM") = "Entitlement earned for the period: " & dlpFrom.Text & " to " & dlpTo.Text & "."
        rzAttend("AD_COMM") = "Entitlement earned for the period: " & rzAttend("AD_DOA") & " to " & dlpTDate.Text & "."
        rzAttend("AD_LDATE") = Date
        rzAttend("AD_LUSER") = glbUserID
        rzAttend("AD_LTIME") = Time$
        rzAttend.Update
        rzAttend.Close
    End If

End Sub

Private Sub vbxTrueGrid1_BeforeRowColChange(Cancel As Integer)
Call SET_UP_MODE
End Sub

Private Sub vbxTrueGrid1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid1_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid1.Tag = "ASC" Then
            vbxTrueGrid1.Tag = "DESC"
        Else
            vbxTrueGrid1.Tag = "ASC"
        End If
        
        If glbtermopen Then
            'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS  from Term_ENTHRS "
            'Ticket #22363 Franks 08/02/2012
            'SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS  FROM Term_ENTHRS "
            'SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
            
            '8.0 - Ticket #22682 - Option to show in Days
            SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS,  "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "

            SQLQ = SQLQ & " FROM Term_ENTHRS "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            'SQLQ = "Select *,HE_PREV+HE_ENTITLE-HE_TAKEN AS OUTS  from HRENTHRS"
            'Ticket #22363 Franks 08/02/2012
            'SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE  HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE  HE_TAKEN END)) AS OUTS FROM HRENTHRS"
            'SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
            
            '8.0 - Ticket #22682 - Option to show in Days
            SQLQ = "SELECT *,((CASE WHEN HE_PREV IS NULL THEN 0 ELSE HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE HE_TAKEN END)) AS OUTS,  "
            'SQLQ = SQLQ & " ((CASE WHEN HE_PREV IS NULL THEN 0 ELSE HE_PREV END) + HE_ENTITLE - (CASE WHEN HE_TAKEN IS NULL THEN 0 ELSE HE_TAKEN END)) AS OUTSDAY  "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_PREV/HE_DHRS END) AS HE_PREVDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_ENTITLE/HE_DHRS END) AS HE_ENTITLEDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE HE_TAKEN/HE_DHRS END) AS HE_TAKENDAY, "
SQLQ = SQLQ & "(CASE WHEN HE_DHRS=0 THEN 0 ELSE ROUND(HE_ENTITLE/HE_DHRS,2)+ROUND(HE_PREV/HE_DHRS,2)-ROUND(HE_TAKEN/HE_DHRS,2) END) AS OUTSDAY "
            
            SQLQ = SQLQ & " FROM HRENTHRS"
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid1.Columns(ColIndex).DataField & " " & vbxTrueGrid1.Tag

        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If
End Sub

Private Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub
