VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmIPFactors 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Incentive Factors"
   ClientHeight    =   10305
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   16500
   ControlBox      =   0   'False
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
   ScaleHeight     =   10305
   ScaleWidth      =   16500
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdROIC 
      Caption         =   "Update ROIC for Year"
      Height          =   375
      Left            =   4320
      TabIndex        =   54
      Top             =   8760
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalesComm 
      Caption         =   "Update Sales Comm for Year"
      Height          =   375
      Left            =   4320
      TabIndex        =   53
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmdCorpFin 
      Caption         =   "Update Corp Fin for Year"
      Height          =   375
      Left            =   4320
      TabIndex        =   52
      Top             =   7800
      Width           =   2655
   End
   Begin VB.CommandButton cmdBUFin 
      Caption         =   "Update BU Fin for Year"
      Height          =   375
      Left            =   4320
      TabIndex        =   51
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton cmdCreatBonusFile 
      Caption         =   "Goto Create Incentive Plan "
      Height          =   375
      Left            =   240
      TabIndex        =   50
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmdCopyToNextYear 
      Caption         =   "Copy To Next Year"
      Height          =   375
      Left            =   240
      TabIndex        =   49
      Top             =   7800
      Width           =   2655
   End
   Begin VB.CommandButton cmdRepeatAllPlants 
      Caption         =   "Repeat for all Plants"
      Height          =   375
      Left            =   240
      TabIndex        =   48
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelDupRec 
      Caption         =   "Delete Record"
      Height          =   375
      Left            =   13440
      TabIndex        =   29
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopyTo 
      Caption         =   "Copy To"
      Height          =   375
      Left            =   12600
      TabIndex        =   25
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmIPFactors.frx":0000
      Height          =   2355
      Left            =   120
      OleObjectBlob   =   "frmIPFactors.frx":0014
      TabIndex        =   22
      Top             =   120
      Width           =   14235
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   21
      Top             =   9645
      Width           =   16500
      _Version        =   65536
      _ExtentX        =   29104
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
         Left            =   6465
         Top             =   180
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
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   450
      Left            =   13560
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   794
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   15285
      TabIndex        =   26
      Tag             =   "00-Section - Code"
      Top             =   8400
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.PictureBox frmDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   120
      ScaleHeight     =   4770
      ScaleWidth      =   12585
      TabIndex        =   0
      Top             =   2520
      Width           =   12585
      Begin MSMask.MaskEdBox MskTPlantObj 
         DataField       =   "IP_T_PLANT_OBJ"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Tag             =   "01-Low Dollars"
         Top             =   2130
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTBUFin 
         DataField       =   "IP_T_BU_FIN"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   6
         Tag             =   "01-MidPoint Dollars"
         Top             =   2490
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTCorpFin 
         DataField       =   "IP_T_CORP_FIN"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   7
         Tag             =   "01-High Dollars"
         Top             =   2850
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFiscalYear 
         DataField       =   "IP_YEAR"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   1
         Tag             =   "01-High Dollars"
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "IP_SECTION"
         DataSource      =   "data1"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Tag             =   "00-Section - Code"
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin MSMask.MaskEdBox MskTSalesInd 
         DataField       =   "IP_T_SALES_IND"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   8
         Tag             =   "10-Percentage of MidPoint"
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "IP_REGION"
         DataSource      =   "data1"
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Tag             =   "00-Region"
         Top             =   840
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "IP_POSTYPE"
         DataSource      =   "data1"
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   2
         Tag             =   "00-Position Type Code"
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "POTY"
         MaxLength       =   10
      End
      Begin MSMask.MaskEdBox MskTSalesComm 
         DataField       =   "IP_T_SALES_COMM"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   9
         Tag             =   "01-High Dollars"
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTCorpObj 
         DataField       =   "IP_T_CORP_OBJ"
         DataSource      =   "data1"
         Height          =   315
         Left            =   1995
         TabIndex        =   10
         Tag             =   "01-High Dollars"
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTROIC 
         DataField       =   "IP_ROIC"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   17
         Tag             =   "01-High Dollars"
         Top             =   4380
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox MskAPlantObj 
         DataField       =   "IP_A_PLANT_OBJ"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   11
         Tag             =   "01-Low Dollars"
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskABUFin 
         DataField       =   "IP_A_BU_FIN"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   12
         Tag             =   "01-MidPoint Dollars"
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskACorpFin 
         DataField       =   "IP_A_CORP_FIN"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   13
         Tag             =   "01-High Dollars"
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskASalesInd 
         DataField       =   "IP_A_SALES_IND"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   14
         Tag             =   "10-Percentage of MidPoint"
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskASalesComm 
         DataField       =   "IP_A_SALES_COMM"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   15
         Tag             =   "01-High Dollars"
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskACorpObj 
         DataField       =   "IP_A_CORP_OBJ"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5595
         TabIndex        =   16
         Tag             =   "01-High Dollars"
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin VB.Line Line3 
         X1              =   3240
         X2              =   3240
         Y1              =   1680
         Y2              =   4320
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7080
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7080
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Ind"
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
         Left            =   3720
         TabIndex        =   47
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Comm"
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
         Left            =   3720
         TabIndex        =   46
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Corp Obj"
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
         Left            =   3720
         TabIndex        =   45
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Obj"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BU Fin"
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
         Left            =   3720
         TabIndex        =   43
         Top             =   2550
         Width           =   480
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Corp Fin"
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
         Left            =   3720
         TabIndex        =   42
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ROIC"
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
         Left            =   3720
         TabIndex        =   41
         Top             =   4380
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Corp Obj"
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
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Comm"
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
         TabIndex        =   39
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BU"
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
         Index           =   24
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjusted By"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   37
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Targets"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblMidPPer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Ind"
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
         TabIndex        =   28
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lblPlant 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant "
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
         TabIndex        =   24
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         TabIndex        =   23
         Top             =   120
         Width           =   330
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Corp Fin"
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
         TabIndex        =   20
         Top             =   2850
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BU Fin"
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
         TabIndex        =   19
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Obj"
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
         TabIndex        =   18
         Top             =   2130
         Width           =   1095
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   9555
      TabIndex        =   33
      Tag             =   "00-Union"
      Top             =   8400
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
      MaxLength       =   15
   End
   Begin MSMask.MaskEdBox MskFiscalYe2 
      DataSource      =   "data1"
      Height          =   315
      Left            =   9870
      TabIndex        =   31
      Tag             =   "01-High Dollars"
      Top             =   7680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###0"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   9555
      TabIndex        =   32
      Tag             =   "00-Position Type Code"
      Top             =   8040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "POTY"
      MaxLength       =   10
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   8280
      TabIndex        =   56
      Top             =   8060
      Width           =   675
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filters:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   55
      Top             =   7320
      Width           =   585
   End
   Begin VB.Label lblUnionFilter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plant"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   35
      Top             =   8400
      Width           =   690
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   34
      Top             =   7680
      Width           =   405
   End
   Begin VB.Label lblDPlant 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Plant "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13680
      TabIndex        =   27
      Top             =   8430
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmIPFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'True DBGrid changed
Option Explicit
Dim fglbRecords%, fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer, fglbNew
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim xID

Private Function chkIPFactors()
Dim SQLQ As String, Msg As String, dd#, PID&, Factor$
chkIPFactors = False
On Error GoTo chkPosEval_Err
'If Trim(clpCode(2)) = "" Then
'    MsgBox "Band is a required field"
'    'cmbBand.SetFocus
'    Exit Function
'End If

'If Val(MskDollars(0)) = 0 Then
'    MsgBox "Low Dollars must be greater than 0 "
'    MskDollars(0).SetFocus
'    Exit Function
'End If
'If Val(MskDollars(1)) = 0 Then
'    MsgBox "MidPoint Dollars must be greater than 0 "
'    MskDollars(1).SetFocus
'    Exit Function
'End If
'If Val(MskDollars(2)) = 0 Then
'    MsgBox "High Dollars must be greater than 0 "
'    MskDollars(2).SetFocus
'    Exit Function
'End If

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYear.Text) = 4 Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If

If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(3).SetFocus
        Exit Function
    End If
Else
    MsgBox lblType.Caption & " is a required field"
    clpCode(3).SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) > 0 Then
    If Len(clpCode(2).Text) > 0 And clpCode(2).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(2).SetFocus
        Exit Function
    End If
'Else
'    MsgBox "Plant is a required field"
'    clpCode(0).SetFocus
'    Exit Function
End If

If Len(clpCode(0).Text) > 0 Then
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(0).SetFocus
        Exit Function
    End If
'Else
'    MsgBox "Plant is a required field"
'    clpCode(0).SetFocus
'    Exit Function
End If

'If modISDupFactor(glbPos$, Factor$, PID&) And fglbNew Then
'    MsgBox "[Band + Market Line + Fiscal Year + Plant] must be unique"
'    'cmbBand.SetFocus
'    Exit Function
'End If
'If Val(MskDollars(1)) < Val(MskDollars(0)) Then
'    MsgBox "MidPoint Dollars must be greater than Low Dollars"
'    MskDollars(1).SetFocus
'    Exit Function
'End If
'
'If Val(MskDollars(2)) < Val(MskDollars(1)) Then
'    MsgBox "High Dollars must be greater than MidPoint Dollars"
'    MskDollars(2).SetFocus
'    Exit Function
'End If


chkIPFactors = True

Exit Function

chkPosEval_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkIPFactors", "HRIP_FACTORS", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub clpCode_Change(Index As Integer)
    If Index = 4 Or Index = 5 Then
        If Not clpCode(5).Caption = "Unassigned" Then
            Call EERetrieve
        End If
    End If
End Sub

Private Sub clpCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbBand_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbBand_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


'Private Sub cmbMarketLine_Change()
''clpCode(3) = Left(cmbMarketLine.Text, 4)
''txtCurrencyIndicator = Left(cmbMarketLine, 2)
'End Sub

'Private Sub cmbMarketLine_click()
''clpCode(3) = Left(cmbMarketLine.Text, 4)
''txtCurrencyIndicator = Left(cmbMarketLine, 2)
'End Sub

'Private Sub cmbMarketLine_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmbMarketLine_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Public Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err
fglbNew = False
'Data1.Recordset.CancelUpdate
Data1.Recordset.Cancel

If Not glbSQL And Not glbOracle Then Call Data1.Refresh
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)  ' reset screen's attributes
'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
'clpCode(2).Enabled = False
'clpCode(3).Enabled = False

Me.vbxTrueGrid.SetFocus
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


Public Sub cmdClose_Click()
Unload frmIPCreateSheet
Unload Me
End Sub



Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, SQLQ, x%, xEmpnbr
Dim xband, xMarketLine, DeleteRight, xJob
'Dim XTB As Recordset

Dim snapAssBand As New ADODB.Recordset
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

Screen.MousePointer = HOURGLASS
''SQLQ = "SELECT SH_EMPNBR,ED_SURNAME,ED_FNAME FROM HR_SALARY_HISTORY,HREMP "
''SQLQ = SQLQ & " WHERE SH_BAND = '" & Data1.Recordset("Band") & "'"
''SQLQ = SQLQ & " AND SH_MARKETLINE = '" & Data1.Recordset("MarketLine") & "'"
''SQLQ = SQLQ & " AND ED_EMPNBR=SH_EMPNBR order by SH_EMPNBR"
''
''If snapAssBand.State <> 0 Then snapAssBand.Close
''snapAssBand.Open SQLQ, gdbAdoIhr001, adOpenKeyset
''Screen.MousePointer = DEFAULT
''
''If Not (snapAssBand.BOF And snapAssBand.EOF) Then
''    x% = 0: xEMPNBR = 0
''    Msg = "This record is in the following employees'" & Chr(10) & "salary history:"
''    While Not snapAssBand.EOF And x% < 10
''        If xEMPNBR <> snapAssBand("sh_EMPNBR") Then
''          Msg = Msg & Chr(10) & snapAssBand("ED_surname") & ", " & snapAssBand("ED_FName") & " -  # " & snapAssBand("sh_EMPNBR")
''          x% = x% + 1
''        End If
''        xEMPNBR = snapAssBand("sh_EMPNBR")
''        snapAssBand.MoveNext
''    Wend
''    Msg = Msg & Chr(10) & "Record will not be deleted."
''    MsgBox Msg
''    Exit Sub
''End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

xID = Data1.Recordset("IP_ID")

'SQLQ = "Delete FROM HRIP_FACTORS "
'SQLQ = SQLQ & "where [band]='" & Trim(clpCode(2)) & "'"
'SQLQ = SQLQ & "and MarketLine='" & Trim(clpCode(3)) & "'"
SQLQ = "DELETE FROM HRIP_FACTORS "
SQLQ = SQLQ & "WHERE IP_ID = " & xID & " "

Data1.Recordset.ActiveConnection.Execute SQLQ
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call Display_Value

Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRIP_FACTORS", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub


Public Sub cmdModify_Click()

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

fglbEditMode% = True

On Error GoTo Mod_Err
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
'clpCode(2).Enabled = False
'clpCode(3).Enabled = False
'MskDollars(0).SetFocus
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
Dim SQLQ As String
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If
'Call SET_UP_MODE
'Call ST_UPD_MODE(True)

On Error GoTo AddN_Err

Data1.Recordset.AddNew
'Call Set_Control("B", Me)
'rsDATA.AddNew

fglbEditMode% = True
fglbNew = True
Call SET_UP_MODE
'cmbBand.Enabled = True
'cmbMarketLine.Enabled = True
'clpCode(2).Enabled = True
'clpCode(3).Enabled = True

MskFiscalYear.SetFocus
'clpCode(3).Text = ""
MskTPlantObj.Text = 0
MskTBUFin.Text = 0
MskTCorpFin.Text = 0
MskTSalesInd.Text = 0
MskTSalesComm.Text = 0
MskTCorpObj.Text = 0
MskAPlantObj.Text = 1
MskABUFin.Text = 1
MskACorpFin.Text = 1
MskASalesInd.Text = 1
MskASalesComm.Text = 1
MskACorpObj.Text = 1
MskTROIC.Text = 0

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRIP_FACTORS", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub


Public Sub cmdOK_Click()
On Error GoTo OK_Err

If Not chkIPFactors() Then Exit Sub
'cmbCurrencyIndicator_setup2 Me
'setMarketLine Me
'If Len(clpCode(0).Text) > 0 Then Data1.Recordset("IP_SECTION") = clpCode(0).Text
If Len(clpCode(3).Text) > 0 Then Data1.Recordset("IP_POSTYPE") = clpCode(3).Text

If Len(MskTPlantObj.Text) = 0 Then MskTPlantObj.Text = 0
If Len(MskTBUFin.Text) = 0 Then MskTBUFin.Text = 0
If Len(MskTCorpFin.Text) = 0 Then MskTCorpFin.Text = 0
If Len(MskTSalesInd.Text) = 0 Then MskTSalesInd.Text = 0
If Len(MskTSalesComm.Text) = 0 Then MskTSalesComm.Text = 0
If Len(MskTCorpObj.Text) = 0 Then MskTCorpObj.Text = 0
If Len(MskAPlantObj.Text) = 0 Then MskAPlantObj.Text = 0
If Len(MskABUFin.Text) = 0 Then MskABUFin.Text = 0
If Len(MskACorpFin.Text) = 0 Then MskACorpFin.Text = 0
If Len(MskASalesInd.Text) = 0 Then MskASalesInd.Text = 0
If Len(MskASalesComm.Text) = 0 Then MskASalesComm.Text = 0
If Len(MskACorpObj.Text) = 0 Then MskACorpObj.Text = 0
If Len(MskTROIC.Text) = 0 Then
    MskTROIC.Text = 0
End If

'Data1.Recordset.UpdateBatch
Data1.Recordset.Update
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
xID = Data1.Recordset("IP_ID")
Data1.Refresh
Data1.Recordset.Find "IP_ID= " & xID

fglbNew = False
fglbEditMode% = False
Call Display_Value
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
'clpCode(2).Enabled = False
'clpCode(3).Enabled = False
Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRIP_FACTORS", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Unload Me


End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub


Private Sub clpCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 5 Then
        If Not clpCode(5).Caption = "Unassigned" Then
            Call EERetrieve
        End If
    End If
End Sub

Private Sub cmdBUFin_Click()
    If Not Data1.Recordset.EOF Then
        glbWFC_IPPopFormName = "UpdateBUFin"
        glbWFC_IncePlanID = Data1.Recordset("IP_ID")
        xExpYears = MskFiscalYear.Text
        frmCheckListView.Show 1
        If glbWFC_IncePlanID = -1 Then
            'Cancel, do nothing
        Else
            'reload data
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdCopyTo_Click()
Dim SQLQ
Dim rsFBand As New ADODB.Recordset
Dim a As Integer, Msg$, INo&, x%

    If Len(MskFiscalYear.Text) > 0 Then
        If Not IsNumeric(MskFiscalYear.Text) Then
            MsgBox "Invalid Fiscal Year."
            MskFiscalYear.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Fiscal Year is a required field"
        MskFiscalYear.SetFocus
        Exit Sub
    End If
    If Len(clpCode(0).Text) > 0 Then
        If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
            MsgBox "If code entered it must be known"
            clpCode(0).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Plant is a required field"
        clpCode(0).SetFocus
        Exit Sub
    End If

    If Len(clpCode(1).Text) > 0 Then
        If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
            MsgBox "If code entered it must be known"
            clpCode(1).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Destination Plant is a required field"
        clpCode(1).SetFocus
        Exit Sub
    End If
    
    If clpCode(0).Text = clpCode(1).Text Then
        MsgBox "Destination Plant Code is equal to From Plant Code"
        clpCode(1).SetFocus
        Exit Sub
    End If

    ''SQLQ = "SELECT * FROM HRIP_FACTORS WHERE BAND = '" & clpCode(2) & "' "
    ''SQLQ = SQLQ & "AND MarketLine = '" & clpCode(3) & "' "
    ''SQLQ = SQLQ & "AND FiscalYear = " & MskFiscalYear & " "
    ''SQLQ = SQLQ & "AND IP_SECTION = '" & clpCode(1) & "' "
    ''rsFBand.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    ''If Not rsFBand.EOF Then
    ''    rsFBand.Close
    ''    MsgBox "Duplicate record found."
    ''    clpCode(1).SetFocus
    ''    Exit Sub
    ''End If



    Msg$ = "Are You Sure You Want To Copy this record? "
    a% = MsgBox(Msg$, 36, "Confirm Copy")
    If a% <> 6 Then
        rsFBand.Close
        Exit Sub
    End If
    rsFBand.AddNew
    'rsFBand("BAND") = clpCode(2)
    'rsFBand("MarketLine") = clpCode(3)
    'If Len(MskDollars(0).Text) > 0 Then rsFBand("LDollars") = MskDollars(0).Text
    'If Len(MskDollars(1).Text) > 0 Then rsFBand("MDollars") = MskDollars(1).Text
    'If Len(MskDollars(2).Text) > 0 Then rsFBand("HDollars") = MskDollars(2).Text
    'If Len(MskMidPointPer.Text) > 0 Then rsFBand("MIDPOINT_PER") = MskMidPointPer.Text
    rsFBand("FiscalYear") = MskFiscalYear
    rsFBand("IP_SECTION") = clpCode(1)
    rsFBand.Update
    xID = rsFBand("IP_ID")
    rsFBand.Close
    Data1.Refresh
    Data1.Recordset.Find "ID= " & xID
    Call Display_Value

End Sub

Private Sub cmdCopyToNextYear_Click()
Dim SQLQ
Dim rsFactors As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim a As Integer, Msg$, INo&, x%
Dim xNextYear
Dim xDelNextYear As Boolean

    If Len(MskFiscalYear.Text) > 0 Then
        If Not IsNumeric(MskFiscalYear.Text) Then
            MsgBox "Invalid Year."
            MskFiscalYear.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Year is a required field"
        MskFiscalYear.SetFocus
        Exit Sub
    End If
    xNextYear = Val(MskFiscalYear.Text) + 1
    
    'check if there is record in new year
    xDelNextYear = False
    SQLQ = "SELECT TOP 1 * FROM HRIP_FACTORS WHERE IP_YEAR = " & xNextYear & " "
    If rsFactors.State <> 0 Then rsFactors.Close
    rsFactors.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Msg$ = ""
    If Not rsFactors.EOF Then
        Msg$ = "Found " & xNextYear & " records of Company Incentive Factors. " & Chr(10)
        Msg$ = Msg$ & "The program will delete all records in " & xNextYear & " first" & Chr(10) & Chr(10)
        xDelNextYear = True
    End If
    rsFactors.Close
    
    Msg$ = Msg$ & "Are you sure you want to copy all records from " & MskFiscalYear.Text & " to " & xNextYear & "? "
    a% = MsgBox(Msg$, 36, "Confirm Copy")
    If a% <> 6 Then
        Exit Sub
    End If
    
    If xDelNextYear Then
        SQLQ = "DELETE FROM HRIP_FACTORS WHERE IP_YEAR = " & xNextYear & " "
        gdbAdoIhr001.Execute SQLQ
    End If
     
    Call WFCIPFactorsToNextYear(MskFiscalYear.Text, xNextYear)
    
    MsgBox "    Finished!    "
    'reload data
    Call Form_Load

End Sub


Private Sub WFCIPFactorsToNextYear(xYear, xNextYear)
Dim rsIPFactors As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer, xTot As Integer

    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xYear & " "
    rsIPFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsIPFactors.EOF Then
        xTot = rsIPFactors.RecordCount
    End If
    I = 0
    
    'the next year records have been deleted so just add the records to next year
    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & xNextYear & " "
    If rsAdd.State <> 0 Then rsAdd.Close
    rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    
    Do While Not rsIPFactors.EOF
        MDIMain.panHelp(0).FloodPercent = Int((I / xTot) * 100)
        I = I + 1
        DoEvents

        rsAdd.AddNew
        rsAdd("IP_YEAR") = xNextYear
        rsAdd("IP_POSTYPE") = rsIPFactors("IP_POSTYPE")
        rsAdd("IP_REGION") = rsIPFactors("IP_REGION")
        rsAdd("IP_SECTION") = rsIPFactors("IP_SECTION")
        rsAdd("IP_T_PLANT_OBJ") = rsIPFactors("IP_T_PLANT_OBJ")
        rsAdd("IP_T_BU_FIN") = rsIPFactors("IP_T_BU_FIN")
        rsAdd("IP_T_CORP_FIN") = rsIPFactors("IP_T_CORP_FIN")
        rsAdd("IP_T_SALES_IND") = rsIPFactors("IP_T_SALES_IND")
        rsAdd("IP_T_SALES_COMM") = rsIPFactors("IP_T_SALES_COMM")
        rsAdd("IP_T_CORP_OBJ") = rsIPFactors("IP_T_CORP_OBJ")
        rsAdd("IP_A_PLANT_OBJ") = rsIPFactors("IP_A_PLANT_OBJ")
        rsAdd("IP_A_BU_FIN") = rsIPFactors("IP_A_BU_FIN")
        rsAdd("IP_A_CORP_FIN") = rsIPFactors("IP_A_CORP_FIN")
        rsAdd("IP_A_SALES_IND") = rsIPFactors("IP_A_SALES_IND")
        rsAdd("IP_A_SALES_COMM") = rsIPFactors("IP_A_SALES_COMM")
        rsAdd("IP_A_CORP_OBJ") = rsIPFactors("IP_A_CORP_OBJ")
        rsAdd("IP_ROIC") = rsIPFactors("IP_ROIC")
        rsAdd("IP_LDATE") = Date
        rsAdd("IP_LTIME") = Time$
        rsAdd("IP_LUSER") = glbUserID
        rsAdd.Update
        rsIPFactors.MoveNext
    Loop
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    Screen.MousePointer = DEFAULT

End Sub


Private Sub cmdCorpFin_Click()
    If Not Data1.Recordset.EOF Then
        glbWFC_IPPopFormName = "UpdateCorpFin"
        glbWFC_IncePlanID = Data1.Recordset("IP_ID")
        xExpYears = MskFiscalYear.Text
        frmCheckListView.Show 1
        If glbWFC_IncePlanID = -1 Then
            'Cancel, do nothing
        Else
            'reload data
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdCreatBonusFile_Click()
    Unload Me
    Load frmIPCreateSheet
    frmIPCreateSheet.ZOrder 0
End Sub

Private Sub cmdDelDupRec_Click()
Dim a As Integer, Msg As String, SQLQ, x%, xEmpnbr
Dim xband, xMarketLine, DeleteRight, xJob
'Dim XTB As Recordset

Dim snapAssBand As New ADODB.Recordset
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

Screen.MousePointer = HOURGLASS

Msg = "This function is only for deleting duplicate record"
Msg = Msg & Chr(10) & "Are You Sure You Want To Delete This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
xID = Data1.Recordset("IP_ID")
SQLQ = "Delete FROM HRIP_FACTORS "
SQLQ = SQLQ & "where [ID]=" & xID & " "
Data1.Recordset.ActiveConnection.Execute SQLQ
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call Display_Value

Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRIP_FACTORS", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdRepeatAllPlants_Click()
    If Len(clpCode(0).Text) = 0 Then
        MsgBox "Can not do 'Repeat for all Plants' if the Plant is blank"
        Exit Sub
    End If
    If Not Data1.Recordset.EOF Then
        glbWFC_IPPopFormName = "WIADivisoinList"
        glbWFC_IncePlanID = Data1.Recordset("IP_ID")
        frmCheckListView.Show 1
        If glbWFC_IncePlanID = -1 Then
            'Cancel, do nothing
        Else
            'reload data
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdROIC_Click()
    If Not Data1.Recordset.EOF Then
        glbWFC_IPPopFormName = "UpdateROIC"
        glbWFC_IncePlanID = Data1.Recordset("IP_ID")
        xExpYears = MskFiscalYear.Text
        frmCheckListView.Show 1
        If glbWFC_IncePlanID = -1 Then
            'Cancel, do nothing
        Else
            'reload data
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdSalesComm_Click()
    If Not Data1.Recordset.EOF Then
        glbWFC_IPPopFormName = "UpdateSalesComm"
        glbWFC_IncePlanID = Data1.Recordset("IP_ID")
        xExpYears = MskFiscalYear.Text
        frmCheckListView.Show 1
        If glbWFC_IncePlanID = -1 Then
            'Cancel, do nothing
        Else
            'reload data
            Call Form_Load
        End If
    End If
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRIP_FACTORS", "SELECT")


End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim x%
glbOnTop = "frmIPFactors"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS
Me.Caption = "Company Incentive Factors"

Data1.ConnectionString = glbAdoIHRDB 'glbAdoIHRWFC
'Data1.RecordSource = "HRIP_FACTORS"



Screen.MousePointer = DEFAULT
x% = EERetrieve()

'Band_AddItem Me
'MarketLine_AddItem Me
'CurrencyIndicator_AddItem Me
fglbNew = False

Call Display_Value

Screen.MousePointer = HOURGLASS
Call INI_Controls(Me)
'Me.vbxTrueGrid.SetFocus


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
Unload frmIPCreateSheet

End Sub

Private Function EERetrieve()
Dim SQLQ
EERetrieve = False
Screen.MousePointer = HOURGLASS
On Error GoTo modGetPosEvalsErr

'SQLQ = "SELECT * FROM HRIP_FACTORS "
SQLQ = "SELECT LTRIM(STR(IP_YEAR)) AS STRYEAR,* FROM HRIP_FACTORS "
SQLQ = SQLQ & "WHERE (1=1) "
If Len(glbPlantCode) > 0 Then
    SQLQ = SQLQ & "AND IP_SECTION = '" & glbPlantCode & "' "
End If
If Len(MskFiscalYe2.Text) > 0 Then
    If IsNumeric(MskFiscalYe2.Text) Then
        SQLQ = SQLQ & "AND IP_YEAR = " & MskFiscalYe2.Text & " "
    End If
End If
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & "AND IP_POSTYPE = '" & clpCode(4).Text & "' "
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ = SQLQ & "AND IP_SECTION = '" & clpCode(5).Text & "' "
End If

SQLQ = SQLQ & "ORDER BY IP_YEAR ,IP_POSTYPE,IP_REGION,IP_SECTION"

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


modGetPosEvalsErr:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Pos Skills", "HRJOBSK", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function





End Function

Private Function modISDupFactor(Pos$, Factor$, PID&)
Dim SQLQ As String
Dim snapEval As New ADODB.Recordset

modISDupFactor = True

On Error GoTo modISDupFactor_Err
Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRIP_FACTORS "
'SQLQ = SQLQ & "where [band]='" & Trim(clpCode(2)) & "'"
'SQLQ = SQLQ & " and MarketLine='" & Trim(clpCode(3)) & "'"
SQLQ = SQLQ & " and IP_SECTION='" & Trim(clpCode(0)) & "'"
SQLQ = SQLQ & " and FiscalYear='" & MskFiscalYear & "'"

snapEval.Open SQLQ, gdbAdoIhr001

If snapEval.BOF And snapEval.EOF Then
    modISDupFactor = False
End If

snapEval.Close
Screen.MousePointer = DEFAULT

Exit Function

modISDupFactor_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub MskDollars_GotFocus(Index As Integer)
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

glbOHSEdit% = TF

fUPMode = TF    ' update mode
frmDetails.Enabled = TF


MskFiscalYear.Enabled = TF
clpCode(0).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
MskTPlantObj.Enabled = TF
MskTBUFin.Enabled = TF
MskTCorpFin.Enabled = TF
MskTSalesInd.Enabled = TF
MskTSalesComm.Enabled = TF
MskTCorpObj.Enabled = TF
MskAPlantObj.Enabled = TF
MskABUFin.Enabled = TF
MskACorpFin.Enabled = TF
MskASalesInd.Enabled = TF
MskASalesComm.Enabled = TF
MskACorpObj.Enabled = TF
MskTROIC.Enabled = TF



'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'vbxTrueGrid.Enabled = FT
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
'    cmdDelete.Enabled = False
'    cmdModify.Enabled = False
'End If

End Sub





Private Sub MskABUFin_GotFocus()
If IsNumeric(MskABUFin.Text) Then
    MskABUFin.Text = MskABUFin.Text * 100
End If
End Sub

Private Sub MskABUFin_LostFocus()
If IsNumeric(MskABUFin.Text) Then
    MskABUFin.Text = MskABUFin.Text / 100
End If
End Sub

Private Sub MskACorpFin_GotFocus()
If IsNumeric(MskACorpFin.Text) Then
    MskACorpFin.Text = MskACorpFin.Text * 100
End If
End Sub

Private Sub MskACorpFin_LostFocus()
If IsNumeric(MskACorpFin.Text) Then
    MskACorpFin.Text = MskACorpFin.Text / 100
End If
End Sub

Private Sub MskACorpObj_GotFocus()
If IsNumeric(MskACorpObj.Text) Then
    MskACorpObj.Text = MskACorpObj.Text * 100
End If
End Sub

Private Sub MskACorpObj_LostFocus()
If IsNumeric(MskACorpObj.Text) Then
    MskACorpObj.Text = MskACorpObj.Text / 100
End If
End Sub

Private Sub MskAPlantObj_GotFocus()
If IsNumeric(MskAPlantObj.Text) Then
    MskAPlantObj.Text = MskAPlantObj.Text * 100
End If
End Sub

Private Sub MskAPlantObj_LostFocus()
If IsNumeric(MskAPlantObj.Text) Then
    MskAPlantObj.Text = MskAPlantObj.Text / 100
End If
End Sub

Private Sub MskASalesComm_GotFocus()
If IsNumeric(MskASalesComm.Text) Then
    MskASalesComm.Text = MskASalesComm.Text * 100
End If
End Sub

Private Sub MskASalesComm_LostFocus()
If IsNumeric(MskASalesComm.Text) Then
    MskASalesComm.Text = MskASalesComm.Text / 100
End If
End Sub

Private Sub MskASalesInd_GotFocus()
If IsNumeric(MskASalesInd.Text) Then
    MskASalesInd.Text = MskASalesInd.Text * 100
End If
End Sub

Private Sub MskASalesInd_LostFocus()
If IsNumeric(MskASalesInd.Text) Then
    MskASalesInd.Text = MskASalesInd.Text / 100
End If
End Sub


Private Sub MskFiscalYe2_Change()
    If Not clpCode(5).Caption = "Unassigned" Then
        Call EERetrieve
    End If
End Sub

Private Sub MskFiscalYe2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not clpCode(5).Caption = "Unassigned" Then
        Call EERetrieve
    End If
End Sub

Private Sub MskTBUFin_GotFocus()
If IsNumeric(MskTBUFin.Text) Then
    MskTBUFin.Text = MskTBUFin.Text * 100
End If
End Sub

Private Sub MskTBUFin_LostFocus()
If IsNumeric(MskTBUFin.Text) Then
    MskTBUFin.Text = MskTBUFin.Text / 100
End If
End Sub

Private Sub MskTCorpFin_GotFocus()
If IsNumeric(MskTCorpFin.Text) Then
    MskTCorpFin.Text = MskTCorpFin.Text * 100
End If
End Sub

Private Sub MskTCorpFin_LostFocus()
If IsNumeric(MskTCorpFin.Text) Then
    MskTCorpFin.Text = MskTCorpFin.Text / 100
End If
End Sub

Private Sub MskTCorpObj_GotFocus()
If IsNumeric(MskTCorpObj.Text) Then
    MskTCorpObj.Text = MskTCorpObj.Text * 100
End If
End Sub

Private Sub MskTCorpObj_LostFocus()
If IsNumeric(MskTCorpObj.Text) Then
    MskTCorpObj.Text = MskTCorpObj.Text / 100
End If
End Sub

Private Sub MskTPlantObj_GotFocus()
If IsNumeric(MskTPlantObj.Text) Then
    MskTPlantObj.Text = MskTPlantObj.Text * 100
End If
End Sub

Private Sub MskTPlantObj_LostFocus()
If IsNumeric(MskTPlantObj.Text) Then
    MskTPlantObj.Text = MskTPlantObj.Text / 100
End If
End Sub


Private Sub MskTSalesComm_GotFocus()
If IsNumeric(MskTSalesComm.Text) Then
    MskTSalesComm.Text = MskTSalesComm.Text * 100
End If
End Sub

Private Sub MskTSalesComm_LostFocus()
If IsNumeric(MskTSalesComm.Text) Then
    MskTSalesComm.Text = MskTSalesComm.Text / 100
End If
End Sub

Private Sub MskTSalesInd_GotFocus()
If IsNumeric(MskTSalesInd.Text) Then
    MskTSalesInd.Text = MskTSalesInd.Text * 100
End If
End Sub

Private Sub MskTSalesInd_LostFocus()
If IsNumeric(MskTSalesInd.Text) Then
    MskTSalesInd.Text = MskTSalesInd.Text / 100
End If
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
        
        SQLQ = "SELECT LTRIM(STR(IP_YEAR)) AS STRYEAR,* FROM HRIP_FACTORS "
        'If Len(glbPlantCode) > 0 Then
        '    SQLQ = SQLQ & "WHERE IP_SECTION = '" & glbPlantCode & "' "
        'End If
        SQLQ = SQLQ & "WHERE (1=1) "
        If Len(glbPlantCode) > 0 Then
            SQLQ = SQLQ & "AND IP_SECTION = '" & glbPlantCode & "' "
        End If
        If Len(clpCode(5).Text) > 0 Then
            SQLQ = SQLQ & "AND IP_SECTION = '" & clpCode(5).Text & "' "
        End If
        If Len(MskFiscalYe2.Text) > 0 Then
            SQLQ = SQLQ & "AND FiscalYear = " & MskFiscalYe2.Text & " "
        End If
        
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    'If cmdOK.Enabled Then
    '    cmdOK.SetFocus
    'Else
    '    cmdClose.SetFocus
    'End If
End If

End Sub



Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call Display_Value
  'cmbBand_SETUP Me
  'cmbCurrencyIndicator_setup2 Me
  'setMarketLine Me
  'MarketLine_Desc Me
  
  
End Sub
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        'rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRIP_FACTORS "
    If Data1.Recordset("IP_ID") > 0 Then
        SQLQ = SQLQ & "WHERE IP_ID = " & Data1.Recordset("IP_ID") & " "
    End If
    SQLQ = SQLQ & "ORDER BY IP_YEAR ,IP_POSTYPE,IP_REGION,IP_SECTION"
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    'rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    'Call Set_Control("R", Me, rsDATA)
    
    'If Not IsNull(rsDATA("IP_POSTYPE")) Then
    '    clpCode(3).Text = rsDATA("IP_POSTYPE")
    'Else
    '    clpCode(3).Text = ""
    'End If
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Get RelateMode() As RelateModeEnum
RelateMode = nothingrelate 'RelatePos
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Inq_SalaryGrids
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


