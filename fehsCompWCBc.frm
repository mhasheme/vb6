VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSCOMPCost 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Company Associated Costs"
   ClientHeight    =   8490
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   11580
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
   ScaleHeight     =   8490
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFirmNo 
      Appearance      =   0  'Flat
      DataField       =   "CA_FIRMNO"
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
      Left            =   7135
      MaxLength       =   10
      TabIndex        =   42
      Tag             =   "00-Firm #"
      Top             =   3680
      Width           =   1350
   End
   Begin VB.TextBox txtLabel10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL10"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   21
      Tag             =   "00-Custom Label 10"
      Top             =   6690
      Width           =   2250
   End
   Begin VB.TextBox txtLabel9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL9"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   19
      Tag             =   "00-Custom Label 9"
      Top             =   6324
      Width           =   2250
   End
   Begin VB.TextBox txtLabel8 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL8"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   17
      Tag             =   "00-Custom Label 8"
      Top             =   5961
      Width           =   2250
   End
   Begin VB.TextBox txtLabel7 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL7"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   15
      Tag             =   "00-Custom Label 7"
      Top             =   5598
      Width           =   2250
   End
   Begin VB.TextBox txtLabel6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL6"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   13
      Tag             =   "00-Custom Label 6"
      Top             =   5235
      Width           =   2250
   End
   Begin VB.TextBox txtLabel5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL5"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "00-Custom Label 5"
      Top             =   4872
      Width           =   2250
   End
   Begin VB.TextBox txtLabel4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL4"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   9
      Tag             =   "00-Custom Label 4"
      Top             =   4509
      Width           =   2250
   End
   Begin VB.TextBox txtLabel3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL3"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   7
      Tag             =   "00-Custom Label 3"
      Top             =   4146
      Width           =   2250
   End
   Begin VB.TextBox txtLabel2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL2"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "00-Custom Label 2"
      Top             =   3783
      Width           =   2250
   End
   Begin VB.TextBox txtLabel1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "CA_LABEL1"
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "00-Custom Label 1"
      Top             =   3420
      Width           =   2250
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsCompWCBc.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "fehsCompWCBc.frx":0014
      TabIndex        =   23
      Top             =   600
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpFromTo 
      DataField       =   "CA_TDATE"
      Height          =   285
      Index           =   1
      Left            =   6825
      TabIndex        =   2
      Tag             =   "42-Cost To Date"
      Top             =   3270
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFromTo 
      DataField       =   "CA_FDATE"
      Height          =   285
      Index           =   0
      Left            =   6825
      TabIndex        =   1
      Tag             =   "42-Cost From Date"
      Top             =   2880
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpSDate 
      DataField       =   "CA_STMTDT"
      Height          =   285
      Left            =   2250
      TabIndex        =   0
      Tag             =   "41-Company Associated Cost Statement Date"
      Top             =   2880
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      TabIndex        =   36
      Top             =   7830
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
         Left            =   6840
         Top             =   240
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CA_LDATE"
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
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7500
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CA_LTIME"
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
      Left            =   4440
      MaxLength       =   25
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7500
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CA_LUSER"
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
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7500
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
         TabIndex        =   41
         Top             =   135
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   160
         Visible         =   0   'False
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
         Left            =   1320
         TabIndex        =   29
         Top             =   135
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medCost1 
      DataField       =   "CA_COST1"
      Height          =   285
      Left            =   2565
      TabIndex        =   4
      Tag             =   "20-Cost for Temporary Compensation"
      Top             =   3420
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost2 
      DataField       =   "CA_COST2"
      Height          =   285
      Left            =   2565
      TabIndex        =   6
      Tag             =   "20-Cost related to Pension"
      Top             =   3783
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost3 
      DataField       =   "CA_COST3"
      Height          =   285
      Left            =   2565
      TabIndex        =   8
      Tag             =   "20-Cost related to Rehabilitation"
      Top             =   4146
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost4 
      DataField       =   "CA_COST4"
      Height          =   285
      Left            =   2565
      TabIndex        =   10
      Tag             =   "20-Non Economic Loss Award"
      Top             =   4509
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost5 
      DataField       =   "CA_COST5"
      Height          =   285
      Left            =   2565
      TabIndex        =   12
      Tag             =   "20-Cost related to Loss of Earning Pension Award"
      Top             =   4872
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost6 
      DataField       =   "CA_COST6"
      Height          =   285
      Left            =   2565
      TabIndex        =   14
      Tag             =   "20-Cost related to Retirement Pension"
      Top             =   5235
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost7 
      DataField       =   "CA_COST7"
      Height          =   285
      Left            =   2565
      TabIndex        =   16
      Tag             =   "20-Cost related to Re-Employment"
      Top             =   5598
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost8 
      DataField       =   "CA_COST8"
      Height          =   285
      Left            =   2565
      TabIndex        =   18
      Tag             =   "20-Cost of Health Care"
      Top             =   5961
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost9 
      DataField       =   "CA_COST9"
      Height          =   285
      Left            =   2565
      TabIndex        =   20
      Tag             =   "20-Survivor Benefit Costs"
      Top             =   6324
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCost10 
      DataField       =   "CA_COST10"
      Height          =   285
      Left            =   2565
      TabIndex        =   22
      Tag             =   "20-Other (user) specified costs"
      Top             =   6690
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFirmNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Firm #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5805
      TabIndex        =   43
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label lblUpdateDate 
      Caption         =   "Updated Date"
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
      Left            =   5880
      TabIndex        =   40
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label lblUpdDateDesc 
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
      Left            =   6960
      TabIndex        =   39
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblUpdateBy 
      Caption         =   "Updated By"
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
      Left            =   2520
      TabIndex        =   38
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblUserDesc 
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
      Left            =   3480
      TabIndex        =   37
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   5805
      TabIndex        =   35
      Top             =   3315
      Width           =   705
   End
   Begin VB.Label lblFDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5805
      TabIndex        =   34
      Top             =   2925
      Width           =   885
   End
   Begin VB.Label lblStatement 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Statement  Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   33
      Top             =   2925
      Width           =   1395
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
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
      Left            =   1710
      TabIndex        =   31
      Top             =   7620
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CA_COMPNO"
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
      Left            =   315
      TabIndex        =   32
      Top             =   7620
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHSCOMPCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X%
Dim fglbNew
Dim wcb() As Variant
Dim fUPMode As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkHSWCBCs()

Dim SQLQ As String, Msg As String, dd&, tdat As Variant

chkHSWCBCs = False

On Error GoTo chkHSWCBCs_Err

If Len(dlpSDate.Text) >= 1 Then
    If Not IsDate(dlpSDate.Text) Then
        MsgBox "Statement Date is not a valid date."
        dlpSDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Statement Date is required."
    dlpSDate.SetFocus
    Exit Function
End If

If Len(dlpFromTo(0).Text) >= 1 Then
    If Not IsDate(dlpFromTo(0).Text) Then
        MsgBox "From date is not a valid date."
        dlpFromTo(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "From date is required."
    dlpFromTo(0).SetFocus
    Exit Function
End If

If Len(dlpFromTo(1).Text) >= 1 Then
    If Not IsDate(dlpFromTo(1).Text) Then
        MsgBox "To date is not a valid date."
        dlpFromTo(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "To date is required."
    dlpFromTo(1).SetFocus
    Exit Function
End If

dd& = DateDiff("d", CVDate(dlpFromTo(0).Text), CVDate(dlpFromTo(1).Text))

If dd& < 0 Then
    MsgBox "From date must be earlier than To date."
    dlpFromTo(0).SetFocus
    Exit Function
End If

dd& = DateDiff("d", CVDate(dlpSDate.Text), CVDate(tdat))

If dd& > 0 Then
    MsgBox "Statement date must be later than File date."
    dlpSDate.SetFocus
    Exit Function
End If

If Len(Trim(txtFirmNo.Text)) = 0 Then
    MsgBox "Firm # is required."
    txtFirmNo.SetFocus
    Exit Function
End If

'If Len(medOther) = 0 Then medOther = 0
Dim Ctrol As Control

Set Ctrol = medCost1: If Not chkNumeric(Ctrol, txtLabel1.Text) Then Exit Function
Set Ctrol = medCost2: If Not chkNumeric(Ctrol, txtLabel2.Text) Then Exit Function
Set Ctrol = medCost3: If Not chkNumeric(Ctrol, txtLabel3.Text) Then Exit Function
Set Ctrol = medCost4: If Not chkNumeric(Ctrol, txtLabel4.Text) Then Exit Function
Set Ctrol = medCost5: If Not chkNumeric(Ctrol, txtLabel5.Text) Then Exit Function
Set Ctrol = medCost6: If Not chkNumeric(Ctrol, txtLabel6.Text) Then Exit Function
Set Ctrol = medCost7: If Not chkNumeric(Ctrol, txtLabel7.Text) Then Exit Function
Set Ctrol = medCost8: If Not chkNumeric(Ctrol, txtLabel8.Text) Then Exit Function
Set Ctrol = medCost9: If Not chkNumeric(Ctrol, txtLabel9.Text) Then Exit Function
Set Ctrol = medCost10: If Not chkNumeric(Ctrol, txtLabel10.Text) Then Exit Function

chkHSWCBCs = True

Exit Function

chkHSWCBCs_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSWCBCs", "HR_OHS_COMPANY_COST", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmbWCB_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbWCB_LostFocus()
'Dim X%
'X% = cmbWCB.ListIndex
'If X% >= 0 Then
'    lblWCBNo.Caption = wcb(X% + 1, 1)
'    lblCase.Caption = wcb(X% + 1, 2)
'End If

End Sub

Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove Binding Control

Call Display_Value
Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_COMPANY_COST", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me

End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, X%

On Error GoTo Del_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "No Records Found"
    Exit Sub
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HROHSCOS", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdIncident_Click()
'frmEHSINCIDENT.Show
'Unload Me
'End Sub

'Private Sub cmdInjLoc_Click()
'frmEHSINJURY.Show
'Unload Me
'End Sub

Sub cmdModify_Click()
Dim X%

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
dlpSDate.SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OHS_COMPANY_COST", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Sub cmdNew_Click()
Dim SQLQ As String

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    Me.vbxTrueGrid.Enabled = False
'    Data1.RecordSource = "HROHSCOS"
'    Data1.Refresh
'    fglbEmptyNew = True
'End If

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)

'If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

'wcb(1, 1) = ""

lblCNum.Caption = "001"
txtLabel1.Text = ""
txtLabel2.Text = ""
txtLabel3.Text = ""
txtLabel4.Text = ""
txtLabel5.Text = ""
txtLabel6.Text = ""
txtLabel7.Text = ""
txtLabel8.Text = ""
txtLabel9.Text = ""
txtLabel10.Text = ""

medCost1 = 0
medCost2 = 0
medCost3 = 0
medCost4 = 0
medCost5 = 0
medCost6 = 0
medCost7 = 0
medCost8 = 0
medCost9 = 0
medCost10 = 0

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_COMPANY_COST", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdOK_Click()
Dim X%

On Error GoTo Add_Err

If Not chkHSWCBCs() Then Exit Sub

rsDATA.Requery

If fglbNew Then rsDATA.AddNew

Call UpdUStats(Me) ' update user's stats (who did it and when)

'If glbtermopen Then
'    'rsDATA!TERM_SEQ = glbTERM_Seq
'    gdbAdoIhr001X.BeginTrans
'    Call Set_Control("U", Me, rsDATA)
'    rsDATA.Update
'    gdbAdoIhr001X.CommitTrans
'    Data1.Refresh
'Else
    gdbAdoIhr001.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
'End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
If NextFormIF(" Company Associated Costs") Then
    Call cmdNew_Click
End If
fglbNew = False
Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_COMPANY_COST", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Company Associated Costs"
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgHSCompCost.rpt"
Me.vbxCrystal.Connect = RptODBC_SQL

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Company Associated Costs"
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgHSCompCost.rpt"
Me.vbxCrystal.Connect = RptODBC_SQL

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
    
End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

'If glbtermopen Then
'    SQLQ = "SELECT * from Term_OHS_COMPANY_COST "
'    SQLQ = SQLQ & "ORDER BY CA_STMTDT, CA_FDATE, CA_TDATE"
    'SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq & " ORDER BY CC_WCBNBR,CC_STMTDT"
'Else
    SQLQ = "SELECT * from HR_OHS_COMPANY_COST "
    SQLQ = SQLQ & "ORDER BY CA_STMTDT, CA_FDATE, CA_TDATE"
    'SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID & " ORDER BY CC_WCBNBR,CC_STMTDT"
'End If
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Exit Function
EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OHS_COMPANY_COST", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function

Private Sub Form_Activate()
glbOnTop = "FRMEHSCOMPCOST"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSCOMPCOST"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

'ReDim wcb(1, 3) 'laura nov 14, 1997
glbOnTop = "FRMEHSCOMPCOST"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

'If Not glbtermopen Then
'    If glbLEE_ID = 0 Then frmEEFIND.Show 1
'    If glbLEE_ID = 0 Then Unload Me: Exit Sub
'Else
'    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
'    If glbTERM_ID = 0 Then Unload Me: Exit Sub
'End If

If EERetrieve() = False Then
'    MsgBox "Sorry, Employee can not be found"
'    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
'Else
    Me.Show
'    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


'If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS
'If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
'    Me.Caption = "WSIB Cost - " & Left$(glbLEE_SName, 5)
'    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
'End If
'lblEENum.Caption = ShowEmpnbr(lblEEID)

Call ST_UPD_MODE(True)
If Not gSec_Upd_HSCost Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
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
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSCOMPCost = Nothing ' carmen may 00
Call NextForm
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

dlpSDate.Enabled = TF
dlpFromTo(0).Enabled = TF
dlpFromTo(1).Enabled = TF
txtFirmNo.Enabled = TF

txtLabel1.Enabled = TF
txtLabel2.Enabled = TF
txtLabel3.Enabled = TF
txtLabel4.Enabled = TF
txtLabel5.Enabled = TF
txtLabel6.Enabled = TF
txtLabel7.Enabled = TF
txtLabel8.Enabled = TF
txtLabel9.Enabled = TF
txtLabel10.Enabled = TF

medCost1.Enabled = TF
medCost2.Enabled = TF
medCost3.Enabled = TF
medCost4.Enabled = TF
medCost5.Enabled = TF
medCost6.Enabled = TF
medCost7.Enabled = TF
medCost8.Enabled = TF
medCost9.Enabled = TF
medCost10.Enabled = TF


'vbxTrueGrid.Enabled = FT

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
End If

End Sub

Private Sub medCost1_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost10_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost3_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost4_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost5_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost6_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost7_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost8_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCost9_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
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
    
    'If glbtermopen Then
    '    SQLQ = "SELECT * from Term_OHS_COMPANY_COST "
        'SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
   ' Else
        SQLQ = "SELECT * from HR_OHS_COMPANY_COST "
        'SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID
    'End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
 '   If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
 '   End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String, X%, WCBN$, WCBN2$

On Error GoTo Tab1_Err
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    'MsgBox "No Records Found"
End If
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OHS_COMPANY_COST", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub


''' Sam add July 2002 * Remove Binding Control
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
        Call SET_UP_MODE
        'Me.cmdModify_Click
        Exit Sub
    End If

    
    'If glbtermopen Then
    '    SQLQ = "SELECT * from Term_OHS_COMPANY_COST "
    '    SQLQ = SQLQ & "WHERE CA_COST_ID = " & Data1.Recordset!CA_COST_ID
    '    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    '    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    'Else
        SQLQ = "SELECT * from HR_OHS_COMPANY_COST "
        SQLQ = SQLQ & "WHERE CA_COST_ID = " & Data1.Recordset!CA_COST_ID
        'SQLQ = SQLQ & "WHERE CC_WCBC_ID = " & Data1.Recordset!CC_WCBC_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    'Me.cmdModify_Click
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
UpdateRight = gSec_Upd_HSCost
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
    frmEHSCOMPCost.Caption = " WSIB Cost Statements - " & Left$(glbLEE_SName, 5)
    frmEHSCOMPCost.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Function chkNumeric(Ctrol As Control, xLabel)
chkNumeric = False
If Len(Ctrol) = 0 Then
    Ctrol = 0
Else
    If Not IsNumeric(Ctrol.Text) Then
        MsgBox xLabel & " must be numeric"
        Ctrol.SetFocus
        Exit Function
    End If
End If
chkNumeric = True
End Function
