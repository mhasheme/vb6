VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMPOSITIONSWFC 
   Appearance      =   0  'Flat
   Caption         =   "Positions Master"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   960
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10695
   ScaleWidth      =   15105
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport vbxCrystal1 
      Left            =   11160
      Top             =   12120
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9840
      Top             =   12120
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   78
      Top             =   9840
      Width           =   15105
      _Version        =   65536
      _ExtentX        =   26644
      _ExtentY        =   1508
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
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
      Begin VB.CommandButton cmdUptSignApprov 
         Appearance      =   0  'Flat
         Caption         =   "Update Signing Approval"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9240
         TabIndex        =   185
         Top             =   60
         Width           =   2295
      End
      Begin VB.CommandButton cmdCopy2AnotheDiv 
         Appearance      =   0  'Flat
         Caption         =   "Copy to Another Division"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   184
         Top             =   60
         Width           =   2295
      End
      Begin VB.CommandButton cmdCopy2AnotherPlant 
         Appearance      =   0  'Flat
         Caption         =   "Copy to Another Plant"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   183
         Top             =   60
         Width           =   2055
      End
      Begin VB.CommandButton cmdImp 
         Appearance      =   0  'Flat
         Caption         =   "Import From Excel File"
         Height          =   375
         Left            =   360
         TabIndex        =   181
         Top             =   60
         Width           =   2055
      End
      Begin VB.CommandButton cmdExp 
         Appearance      =   0  'Flat
         Caption         =   "Export Into Excel File"
         Height          =   375
         Left            =   2520
         TabIndex        =   182
         Top             =   60
         Width           =   2055
      End
      Begin VB.CommandButton cmdAttachJobFiles 
         Appearance      =   0  'Flat
         Caption         =   "&Job Files..."
         Height          =   495
         Left            =   2280
         TabIndex        =   133
         Tag             =   "Attach Files related to this Job"
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.HScrollBar scrHScroll 
         Height          =   300
         LargeChange     =   25
         Left            =   0
         Max             =   50
         SmallChange     =   4
         TabIndex        =   146
         Top             =   520
         Width           =   11535
      End
      Begin VB.CommandButton cmdCountPos 
         Appearance      =   0  'Flat
         Caption         =   "&Count Positions + Total Points"
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   11280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Position Master"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   7305
      LargeChange     =   315
      Left            =   11520
      Max             =   100
      SmallChange     =   315
      TabIndex        =   80
      Top             =   2220
      Width           =   300
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmpostnwfc.frx":0000
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "fxmpostnwfc.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   11475
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   0
      TabIndex        =   85
      Top             =   1860
      Width           =   11475
      Begin VB.TextBox txtUserDef1 
         Appearance      =   0  'Flat
         DataField       =   "JB_USERDEF1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7600
         MaxLength       =   25
         TabIndex        =   21
         Tag             =   "00-User Defined 1 "
         Top             =   3360
         Width           =   1545
      End
      Begin VB.TextBox medUserDef2 
         Appearance      =   0  'Flat
         DataField       =   "JB_USERDEF2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7600
         TabIndex        =   23
         Tag             =   "00-User Defined 1 "
         Top             =   3720
         Width           =   1545
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_BAND"
         Height          =   285
         Index           =   6
         Left            =   1425
         TabIndex        =   14
         Tag             =   "00-Band - Code"
         Top             =   2355
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin VB.Frame frmWFCDIV 
         Height          =   330
         Left            =   1425
         TabIndex        =   173
         Top             =   1350
         Width           =   4215
         Begin VB.TextBox txtJobCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "JB_JOBCODE"
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
            Left            =   320
            MaxLength       =   25
            TabIndex        =   8
            Tag             =   "01-Job Code"
            Top             =   0
            Width           =   1110
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "fxmpostnwfc.frx":B9DC
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblJobCodeDesc 
            Caption         =   "Unassigned"
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
            Height          =   255
            Left            =   1560
            TabIndex        =   174
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.TextBox txtMercerNo 
         Appearance      =   0  'Flat
         DataField       =   "JB_MERCER_NO"
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
         Left            =   7600
         MaxLength       =   25
         TabIndex        =   7
         Tag             =   "00-Mercer Code"
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CheckBox chkHideInactive 
         Caption         =   "Hide Inactive Positions"
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
         Left            =   9480
         TabIndex        =   149
         Top             =   48
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.Frame fraGrid1 
         Appearance      =   0  'Flat
         Caption         =   "Grid Steps"
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
         ForeColor       =   &H80000008&
         Height          =   6735
         Left            =   11040
         TabIndex        =   132
         Top             =   6600
         Width           =   2250
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S10A"
            Height          =   285
            Index           =   30
            Left            =   480
            TabIndex        =   66
            Tag             =   "20-Grid Scales for position"
            Top             =   3111
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S9A"
            Height          =   285
            Index           =   29
            Left            =   480
            TabIndex        =   65
            Tag             =   "20-Grid Scales for position"
            Top             =   2792
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S8A"
            Height          =   285
            Index           =   28
            Left            =   480
            TabIndex        =   64
            Tag             =   "20-Grid Scales for position"
            Top             =   2473
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S7A"
            Height          =   285
            Index           =   27
            Left            =   480
            TabIndex        =   63
            Tag             =   "20-Grid Scales for position"
            Top             =   2154
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S6A"
            Height          =   285
            Index           =   26
            Left            =   480
            TabIndex        =   62
            Tag             =   "20-Grid Scales for position"
            Top             =   1835
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S5A"
            Height          =   285
            Index           =   25
            Left            =   480
            TabIndex        =   61
            Tag             =   "20-Grid Scales for position"
            Top             =   1516
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S4A"
            Height          =   285
            Index           =   24
            Left            =   480
            TabIndex        =   60
            Tag             =   "20-Grid Scales for position"
            Top             =   1197
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S3A"
            Height          =   285
            Index           =   23
            Left            =   480
            TabIndex        =   59
            Tag             =   "20-Grid Scales for position"
            Top             =   878
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S2A"
            Height          =   285
            Index           =   22
            Left            =   480
            TabIndex        =   58
            Tag             =   "20-Grid Scales for position"
            Top             =   559
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S1A"
            Height          =   285
            Index           =   21
            Left            =   480
            TabIndex        =   57
            Tag             =   "21-Grid Scales for position"
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S11A"
            Height          =   285
            Index           =   31
            Left            =   480
            TabIndex        =   67
            Tag             =   "20-Grid Scales for position"
            Top             =   3430
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S12A"
            Height          =   285
            Index           =   32
            Left            =   480
            TabIndex        =   68
            Tag             =   "20-Grid Scales for position"
            Top             =   3749
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S13A"
            Height          =   285
            Index           =   33
            Left            =   480
            TabIndex        =   69
            Tag             =   "20-Grid Scales for position"
            Top             =   4068
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S14A"
            Height          =   285
            Index           =   34
            Left            =   480
            TabIndex        =   70
            Tag             =   "20-Grid Scales for position"
            Top             =   4387
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S15A"
            Height          =   285
            Index           =   35
            Left            =   480
            TabIndex        =   71
            Tag             =   "20-Grid Scales for position"
            Top             =   4706
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S16A"
            Height          =   285
            Index           =   36
            Left            =   480
            TabIndex        =   72
            Tag             =   "20-Grid Scales for position"
            Top             =   5025
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S17A"
            Height          =   285
            Index           =   37
            Left            =   480
            TabIndex        =   73
            Tag             =   "20-Grid Scales for position"
            Top             =   5344
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S18A"
            Height          =   285
            Index           =   38
            Left            =   480
            TabIndex        =   74
            Tag             =   "20-Grid Scales for position"
            Top             =   5663
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S19A"
            Height          =   285
            Index           =   39
            Left            =   480
            TabIndex        =   75
            Tag             =   "20-Grid Scales for position"
            Top             =   5982
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S20A"
            Height          =   285
            Index           =   40
            Left            =   480
            TabIndex        =   76
            Tag             =   "20-Grid Scales for position"
            Top             =   6315
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "16"
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
            Index           =   35
            Left            =   210
            TabIndex        =   167
            Top             =   5070
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "17"
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
            Index           =   36
            Left            =   210
            TabIndex        =   166
            Top             =   5389
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "18"
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
            Index           =   37
            Left            =   210
            TabIndex        =   165
            Top             =   5708
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "19"
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
            Index           =   38
            Left            =   210
            TabIndex        =   164
            Top             =   6027
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "20"
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
            Index           =   39
            Left            =   210
            TabIndex        =   163
            Top             =   6360
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "15"
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
            Index           =   34
            Left            =   210
            TabIndex        =   157
            Top             =   4751
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "14"
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
            Index           =   33
            Left            =   210
            TabIndex        =   156
            Top             =   4432
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "13"
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
            Index           =   32
            Left            =   210
            TabIndex        =   155
            Top             =   4113
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "12"
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
            Index           =   31
            Left            =   210
            TabIndex        =   154
            Top             =   3794
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "11"
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
            Index           =   30
            Left            =   210
            TabIndex        =   144
            Top             =   3475
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "8"
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
            Index           =   27
            Left            =   300
            TabIndex        =   143
            Top             =   2518
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "7"
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
            Index           =   26
            Left            =   300
            TabIndex        =   142
            Top             =   2199
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "10"
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
            Index           =   29
            Left            =   210
            TabIndex        =   141
            Top             =   3156
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "9"
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
            Index           =   28
            Left            =   300
            TabIndex        =   140
            Top             =   2837
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Left            =   300
            TabIndex        =   139
            Top             =   1561
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Index           =   23
            Left            =   300
            TabIndex        =   138
            Top             =   1242
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Index           =   25
            Left            =   300
            TabIndex        =   137
            Top             =   1880
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Index           =   22
            Left            =   300
            TabIndex        =   136
            Top             =   923
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Index           =   21
            Left            =   300
            TabIndex        =   135
            Top             =   604
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Index           =   20
            Left            =   300
            TabIndex        =   134
            Top             =   285
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   7740
         TabIndex        =   130
         Top             =   80
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbMidPoint 
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
         Left            =   7560
         TabIndex        =   31
         Tag             =   "01-Mid Point Grid Step number"
         Text            =   "cmbMidPoint"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtMidPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "JB_MIDPOINT"
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
         Left            =   7950
         MaxLength       =   2
         TabIndex        =   98
         Top             =   6960
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtNoPos 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         DataField       =   "JB_NBRPOS"
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   32
         Tag             =   "10-Number of positions that exist for this job"
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Frame fraGrid 
         Appearance      =   0  'Flat
         Caption         =   "Grid Steps"
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
         Height          =   6735
         Left            =   9600
         TabIndex        =   86
         Top             =   6960
         Width           =   2250
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S10"
            Height          =   285
            Index           =   10
            Left            =   480
            TabIndex        =   46
            Tag             =   "20-Grid Scales for position"
            Top             =   3111
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S9"
            Height          =   285
            Index           =   9
            Left            =   480
            TabIndex        =   45
            Tag             =   "20-Grid Scales for position"
            Top             =   2792
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S8"
            Height          =   285
            Index           =   8
            Left            =   480
            TabIndex        =   44
            Tag             =   "20-Grid Scales for position"
            Top             =   2473
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S7"
            Height          =   285
            Index           =   7
            Left            =   480
            TabIndex        =   43
            Tag             =   "20-Grid Scales for position"
            Top             =   2154
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S6"
            Height          =   285
            Index           =   6
            Left            =   480
            TabIndex        =   42
            Tag             =   "20-Grid Scales for position"
            Top             =   1835
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S5"
            Height          =   285
            Index           =   5
            Left            =   480
            TabIndex        =   41
            Tag             =   "20-Grid Scales for position"
            Top             =   1516
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S4"
            Height          =   285
            Index           =   4
            Left            =   480
            TabIndex        =   40
            Tag             =   "20-Grid Scales for position"
            Top             =   1197
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S3"
            Height          =   285
            Index           =   3
            Left            =   480
            TabIndex        =   39
            Tag             =   "20-Grid Scales for position"
            Top             =   878
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S2"
            Height          =   285
            Index           =   2
            Left            =   480
            TabIndex        =   38
            Tag             =   "20-Grid Scales for position"
            Top             =   559
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S1"
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   37
            Tag             =   "21-Grid Scales for position"
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S11"
            Height          =   285
            Index           =   11
            Left            =   480
            TabIndex        =   47
            Tag             =   "20-Grid Scales for position"
            Top             =   3430
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S12"
            Height          =   285
            Index           =   12
            Left            =   480
            TabIndex        =   48
            Tag             =   "20-Grid Scales for position"
            Top             =   3749
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S13"
            Height          =   285
            Index           =   13
            Left            =   480
            TabIndex        =   49
            Tag             =   "20-Grid Scales for position"
            Top             =   4068
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S14"
            Height          =   285
            Index           =   14
            Left            =   480
            TabIndex        =   50
            Tag             =   "20-Grid Scales for position"
            Top             =   4387
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S15"
            Height          =   285
            Index           =   15
            Left            =   480
            TabIndex        =   51
            Tag             =   "20-Grid Scales for position"
            Top             =   4706
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S16"
            Height          =   285
            Index           =   16
            Left            =   480
            TabIndex        =   52
            Tag             =   "20-Grid Scales for position"
            Top             =   5025
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S17"
            Height          =   285
            Index           =   17
            Left            =   480
            TabIndex        =   53
            Tag             =   "20-Grid Scales for position"
            Top             =   5344
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S18"
            Height          =   285
            Index           =   18
            Left            =   480
            TabIndex        =   54
            Tag             =   "20-Grid Scales for position"
            Top             =   5663
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S19"
            Height          =   285
            Index           =   19
            Left            =   480
            TabIndex        =   55
            Tag             =   "20-Grid Scales for position"
            Top             =   5982
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S20"
            Height          =   285
            Index           =   20
            Left            =   480
            TabIndex        =   56
            Tag             =   "20-Grid Scales for position"
            Top             =   6315
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "16"
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
            Index           =   15
            Left            =   195
            TabIndex        =   162
            Top             =   5070
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "17"
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
            Index           =   16
            Left            =   195
            TabIndex        =   161
            Top             =   5389
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "18"
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
            Index           =   17
            Left            =   195
            TabIndex        =   160
            Top             =   5708
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "19"
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
            Index           =   18
            Left            =   195
            TabIndex        =   159
            Top             =   6027
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "20"
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
            Left            =   195
            TabIndex        =   158
            Top             =   6360
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "15"
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
            Index           =   14
            Left            =   195
            TabIndex        =   153
            Top             =   4751
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "14"
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
            Left            =   195
            TabIndex        =   152
            Top             =   4432
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "13"
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
            Left            =   195
            TabIndex        =   151
            Top             =   4113
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "12"
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
            Left            =   195
            TabIndex        =   150
            Top             =   3794
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "11"
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
            Left            =   195
            TabIndex        =   97
            Top             =   3475
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "8"
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
            Left            =   285
            TabIndex        =   96
            Top             =   2518
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "7"
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
            Left            =   285
            TabIndex        =   95
            Top             =   2199
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "10"
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
            Left            =   195
            TabIndex        =   94
            Top             =   3156
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "9"
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
            Left            =   285
            TabIndex        =   93
            Top             =   2837
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Left            =   285
            TabIndex        =   92
            Top             =   1561
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Left            =   285
            TabIndex        =   91
            Top             =   1242
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Left            =   285
            TabIndex        =   90
            Top             =   1880
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Left            =   285
            TabIndex        =   89
            Top             =   923
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Left            =   285
            TabIndex        =   88
            Top             =   604
            Width           =   90
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Left            =   285
            TabIndex        =   87
            Top             =   285
            Width           =   90
         End
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         DataField       =   "JB_CODE"
         Height          =   315
         Left            =   1740
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "01-Position Code (Unique)"
         Top             =   10
         Width           =   1305
      End
      Begin VB.TextBox txtPosDescr 
         Appearance      =   0  'Flat
         DataField       =   "JB_DESCR"
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
         Left            =   1740
         MaxLength       =   100
         TabIndex        =   2
         Tag             =   "01-Position Description"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtPosDescr2 
         Appearance      =   0  'Flat
         DataField       =   "JB_DESCR2"
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
         Left            =   7600
         MaxLength       =   100
         TabIndex        =   3
         Tag             =   "00-Position Alternate Description"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox medFTENum 
         Appearance      =   0  'Flat
         DataField       =   "JB_FTENUM"
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
         Left            =   1620
         TabIndex        =   27
         Tag             =   "10-Number of FTE "
         Top             =   5580
         Width           =   1215
      End
      Begin VB.TextBox medFTEHrs 
         Appearance      =   0  'Flat
         DataField       =   "JB_FTEHRS"
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
         Left            =   1620
         TabIndex        =   28
         Tag             =   "10-FTE Hours/Year"
         Top             =   5880
         Width           =   1215
      End
      Begin VB.ComboBox comPayPer 
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
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "01- Grid Steps - Annual or Hourly"
         Top             =   2355
         Width           =   2730
      End
      Begin INFOHR_Controls.CodeLookup clpLGroup 
         DataField       =   "JB_LOCGROUP"
         Height          =   285
         Left            =   1425
         TabIndex        =   35
         Tag             =   "Location Group"
         Top             =   7650
         Visible         =   0   'False
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBLC"
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU3"
         Height          =   285
         Index           =   2
         Left            =   1425
         TabIndex        =   22
         Tag             =   "00-Enter Position Code"
         Top             =   3615
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU2"
         Height          =   285
         Index           =   1
         Left            =   1425
         TabIndex        =   20
         Tag             =   "00-Enter Position Code"
         Top             =   3300
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU"
         Height          =   285
         Index           =   0
         Left            =   1425
         TabIndex        =   18
         Tag             =   "00-Enter Position Code"
         Top             =   2985
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpNationalClass 
         DataField       =   "JB_FEDGRP"
         Height          =   285
         Left            =   7290
         TabIndex        =   19
         Tag             =   "00-National Occupation Classification -Code"
         Top             =   3030
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   6
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_WF2"
         Height          =   285
         Index           =   5
         Left            =   1425
         TabIndex        =   34
         Tag             =   "01-WF2 Code"
         Top             =   7320
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "LNWF"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_WF1"
         Height          =   285
         Index           =   4
         Left            =   1425
         TabIndex        =   33
         Tag             =   "01-WF1 Code"
         Top             =   6990
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "LNWF"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_ORG"
         Height          =   285
         Index           =   3
         Left            =   7290
         TabIndex        =   17
         Tag             =   "00-Union - Code "
         Top             =   2700
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_GRPCD"
         Height          =   285
         Index           =   2
         Left            =   7290
         TabIndex        =   11
         Tag             =   "01-Position Group Code "
         Top             =   1680
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_STATUS"
         Height          =   285
         Index           =   1
         Left            =   1425
         TabIndex        =   10
         Tag             =   "01-Position Status - Code "
         Top             =   1680
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBST"
         MaxLength       =   6
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         Height          =   315
         Left            =   2160
         TabIndex        =   36
         Top             =   7980
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         ShowUnassigned  =   1
         TABLName        =   "JBGD"
         Object.Height          =   315
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JB_DHRS"
         Height          =   285
         Left            =   1620
         TabIndex        =   29
         Tag             =   "10-Usual working hours per day"
         Top             =   6180
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_MACHINE_NUM"
         Height          =   285
         Index           =   0
         Left            =   1425
         TabIndex        =   147
         Tag             =   "00-Machine #"
         Top             =   6630
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   15
         LookupType      =   12
      End
      Begin MSMask.MaskEdBox medPoints 
         DataField       =   "JB_POINTS"
         Height          =   285
         Left            =   7600
         TabIndex        =   9
         Tag             =   "10-Total Points"
         Top             =   1350
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_REGION"
         Height          =   285
         Index           =   9
         Left            =   7290
         TabIndex        =   5
         Tag             =   "00-Region"
         Top             =   690
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_SECTION"
         Height          =   285
         Index           =   7
         Left            =   1425
         TabIndex        =   6
         Tag             =   "00-Section - Code"
         Top             =   1020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "JB_DIV"
         Height          =   285
         Left            =   1425
         TabIndex        =   4
         Tag             =   "00-Specific Division Desired"
         Top             =   690
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   0
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_MARKETLINE"
         Height          =   285
         Index           =   8
         Left            =   1425
         TabIndex        =   16
         Tag             =   "00-Market Line - Code"
         Top             =   2670
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFML"
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU4"
         Height          =   285
         Index           =   3
         Left            =   1425
         TabIndex        =   24
         Tag             =   "00-Enter Position Code"
         Top             =   3930
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_LEVEL"
         Height          =   285
         Index           =   10
         Left            =   7290
         TabIndex        =   15
         Tag             =   "01-Job Level Code "
         Top             =   2370
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBLE"
      End
      Begin INFOHR_Controls.DateLookup dlpSDATE 
         DataField       =   "JB_SDATE"
         Height          =   285
         Left            =   1425
         TabIndex        =   12
         Tag             =   "41-Start Date"
         Top             =   2020
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1465
      End
      Begin INFOHR_Controls.DateLookup dlpEDATE 
         DataField       =   "JB_EDATE"
         Height          =   285
         Left            =   7290
         TabIndex        =   13
         Tag             =   "41-End Date"
         Top             =   2020
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1465
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_POSTYPE"
         Height          =   285
         Index           =   11
         Left            =   7290
         TabIndex        =   25
         Tag             =   "00-Position Type Code"
         Top             =   4080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "POTY"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   12
         Left            =   1425
         TabIndex        =   194
         Tag             =   "00-Union"
         Top             =   5040
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
         MaxLength       =   15
      End
      Begin MSMask.MaskEdBox MskApprLimit 
         DataField       =   "JB_APPR_LIMIT"
         Height          =   315
         Left            =   7600
         TabIndex        =   26
         Tag             =   "01-High Dollars"
         Top             =   4410
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblApprLimit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Signing Approval Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   197
         Top             =   4410
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   196
         Top             =   5070
         Width           =   495
      End
      Begin VB.Label lblUnionFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   195
         Top             =   5070
         Width           =   690
      End
      Begin VB.Label lblUptDate 
         Caption         =   "lblUptDate"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         TabIndex        =   193
         Top             =   4600
         Width           =   2415
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   192
         Top             =   4300
         Width           =   1095
      End
      Begin VB.Label lblUserDesc 
         Caption         =   "lblUserDesc"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         TabIndex        =   191
         Top             =   4300
         Width           =   2415
      End
      Begin VB.Label lblUpdateD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   190
         Top             =   4600
         Width           =   1215
      End
      Begin VB.Label lblPosType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   189
         Top             =   4080
         Width           =   1140
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   5760
         TabIndex        =   188
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   30
         TabIndex        =   187
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label lblMissingBudPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This Position Master is missing a Budgeted Position Master record."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5760
         TabIndex        =   186
         Top             =   4920
         Visible         =   0   'False
         Width           =   5670
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   180
         Top             =   2370
         Width           =   1215
      End
      Begin VB.Label lblUserDef1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position User Defined 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   179
         Top             =   3390
         Width           =   1680
      End
      Begin VB.Label lblUserDef2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position User Defined 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   178
         Top             =   3750
         Width           =   1680
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   30
         TabIndex        =   177
         Top             =   3960
         Width           =   930
      End
      Begin VB.Label lblMarketLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   30
         TabIndex        =   176
         Top             =   2685
         Width           =   1065
      End
      Begin VB.Label lblJobCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   175
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label lbltMercerNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mercer Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   172
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   30
         TabIndex        =   171
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblPlant 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   170
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   169
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblPosDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Description"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   30
         TabIndex        =   168
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lblMachine 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   148
         Top             =   7050
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -90
         TabIndex        =   145
         Top             =   6225
         Width           =   930
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   7360
         Picture         =   "fxmpostnwfc.frx":BB26
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   7360
         Picture         =   "fxmpostnwfc.frx":BC70
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Caption         =   "Position Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5760
         TabIndex        =   131
         Top             =   90
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label txtLambtonJob 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   7860
         TabIndex        =   129
         Top             =   7380
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6315
         TabIndex        =   128
         Top             =   7410
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblGridC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Category (Default)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   127
         Top             =   8040
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label lblJobID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6000
         TabIndex        =   126
         Top             =   6960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Status"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   30
         TabIndex        =   125
         Top             =   1695
         Width           =   1245
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Group"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   124
         Top             =   1695
         Width           =   1230
      End
      Begin VB.Label lblSalary 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Grid"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   123
         Top             =   2370
         Width           =   945
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   122
         Top             =   2985
         Width           =   1290
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   121
         Top             =   2700
         Width           =   465
      End
      Begin VB.Label lblProv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.C. Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   120
         Top             =   3045
         Width           =   960
      End
      Begin VB.Label lblNoPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Positions"
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
         Left            =   -90
         TabIndex        =   119
         Top             =   6480
         Width           =   1410
      End
      Begin VB.Label lblFTENum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -90
         TabIndex        =   118
         Top             =   5595
         Width           =   540
      End
      Begin VB.Label lblFTEHrs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE Hours/Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -90
         TabIndex        =   117
         Top             =   5910
         Width           =   1395
      End
      Begin VB.Label lblTotalPoints 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Points"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   116
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   115
         Top             =   1365
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "JB_SALCD"
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
         TabIndex        =   114
         Top             =   7080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPosFilled 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Positions Filled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3150
         TabIndex        =   113
         Top             =   6480
         Width           =   1050
      End
      Begin VB.Label lblTotNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3165
         TabIndex        =   112
         Top             =   5595
         Width           =   1035
      End
      Begin VB.Label lblTotHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE Hours/Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3165
         TabIndex        =   111
         Top             =   5940
         Width           =   1545
      End
      Begin VB.Label lblCountWarn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Warning # Positions < Positions Filled"
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
         Left            =   8520
         TabIndex        =   110
         Top             =   1395
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label lblPosFiled 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_NBRFIL"
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
         Left            =   4980
         TabIndex        =   109
         Top             =   6465
         Width           =   90
      End
      Begin VB.Label lblFTETotNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_FTETOTNU"
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
         Left            =   5100
         TabIndex        =   108
         Top             =   3915
         Width           =   90
      End
      Begin VB.Label lblFTETotHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_FTETOTHR"
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
         Left            =   4920
         TabIndex        =   107
         Top             =   5520
         Width           =   90
      End
      Begin VB.Label lblComRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mid-Point"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   106
         Top             =   6720
         Width           =   840
      End
      Begin VB.Label lblPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   77
         Top             =   40
         Width           =   1410
      End
      Begin VB.Label lblPosAlter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Alternate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   105
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblBand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Band"
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
         Left            =   30
         TabIndex        =   104
         Top             =   2370
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   30
         TabIndex        =   103
         Top             =   3615
         Width           =   930
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   102
         Top             =   3315
         Width           =   930
      End
      Begin VB.Label lblWF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "WF1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   101
         Top             =   6690
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblWF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "WF2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   100
         Top             =   7380
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblLGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location Group "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   150
         TabIndex        =   99
         Top             =   7680
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LUSER"
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
      Left            =   8040
      MaxLength       =   25
      TabIndex        =   84
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LTIME"
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
      Left            =   6840
      MaxLength       =   25
      TabIndex        =   83
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LDATE"
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
      Left            =   5760
      MaxLength       =   25
      TabIndex        =   82
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   7140
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "JB_COMPNO"
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
      Height          =   285
      Left            =   8760
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmMPOSITIONSWFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dynSH_Job1 As New ADODB.Recordset
Dim fglbEditMode%
Dim GridLev(0 To 10)
Dim flagLoad As Integer  'carmen may 2000
Dim IfGridStepChange As Boolean
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'Dim GridStep(11, 2) As Variant
'Dim GridStep2(11, 2) As Variant
'Dim GridStep(15, 2) As Variant
'Dim GridStep2(15, 2) As Variant
Dim GridStep(20, 2) As Variant
Dim GridStep2(20, 2) As Variant

Dim fglbCOMPA#, fglbGRADE$
Dim dblOSalary, dblNewSalary
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, empNo&, dblWHours#, OTOTAL
Dim oPayP, NPayp, OJOB1, OSalCD, oGrade
Dim EmpChgErrors, strEmpEffError As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim lstMidPoint
Dim fglbNew As Boolean
Dim oJobCode, oJobDesc, oJobUnion, oStatus
Dim SkipResetGridStep As Boolean
Dim xStep
Dim oPayrollID
Dim fglbDhrs
Dim strEMPLIST As String

Private Function chkPositions()
Dim X, JobID, Job, intLastNonZero
' unique taken cared of in data control/validate
Dim SQLQ As String, Msg As String

chkPositions = False

On Error GoTo chkPositions_Err

If Len(Trim(txtPosition)) <= 0 Then
    MsgBox "Position Code is a required field"
    txtPosition.SetFocus
    Exit Function
End If
If glbMultiGrid And fglbNew Then
    If Len(clpGrid) = 0 Then
        MsgBox lStr("Default Grid Category is required field")
        clpGrid.SetFocus
    End If
End If
If fglbNew = True Then
    JobID = 0
Else
    JobID = CLng(Val(lblJobID))
End If
Job = txtPosition
If glbWFC Then 'Ticket #25911 Franks 10/01/2014
    If Len(clpCode(9).Text) = 0 Then
        MsgBox lStr("Region") & " is a required field"
        clpCode(9).SetFocus
        Exit Function
    End If
    If Len(clpDiv.Text) = 0 Then
        MsgBox lStr("Division") & " is a required field"
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(7).Text) = 0 Then
        MsgBox lStr("Section") & " is a required field"
        clpCode(7).SetFocus
        Exit Function
    End If
    If Not modJobSectionUnique(Job, JobID, clpCode(7).Text) Then
        MsgBox "[Job Code + " & lStr("Section") & "] is not unique"
        txtPosition.Enabled = True
        txtPosition.SetFocus
        Exit Function
    End If
    If Len(txtJobCode.Text) = 0 Then
        MsgBox "Job Code is a required field"
        txtJobCode.SetFocus
        Exit Function
    End If
    'If clpCode(10).Visible Then
    '    If Len(clpCode(10).Text) = 0 Then
    '        MsgBox lStr("Position Level") & " is a required field"
    '        clpCode(10).SetFocus
    '        Exit Function
    '    End If
    'End If
    If clpCode(6).Visible Then
        If Len(clpCode(6).Text) = 0 Then
            MsgBox "Band is a required field"
            clpCode(6).SetFocus
            Exit Function
        End If
        If Len(clpCode(6).Text) > 0 Then
            If clpCode(6).Caption = "Unassigned" Then
                MsgBox "Band code must be valid"
                clpCode(6).SetFocus
                Exit Function
            End If
        End If
    End If
    'Ticket #29183 Franks 09/12/2016 - begin
    If Len(clpCode(3).Text) = 0 Then
        MsgBox lStr("Union") & " is a required field"
        clpCode(3).SetFocus
        Exit Function
    End If
    
    If txtJobCode.Text = "IND000" Then 'Ticket #30313 Franks 07/03/2017
    Else
        If Len(clpNationalClass.Text) = 0 Then
            MsgBox "N.O.C. Code is a required field"
            clpNationalClass.SetFocus
            Exit Function
        End If
        If Len(clpCode(11).Text) = 0 Then
            MsgBox lStr("Position Type") & " is a required field"
            clpCode(11).SetFocus
            Exit Function
        End If
    End If
    'Ticket #29183 Franks 09/12/2016 - end
Else
    If Not modJobUnique(Job, JobID) Then
        MsgBox "Job Code is not unique"
        txtPosition.Enabled = True
        txtPosition.SetFocus
        Exit Function
    End If
End If

If Len(Trim(txtPosDescr)) = 0 Then
    MsgBox "Position Description is a required field"
    txtPosDescr.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #17490
    If Len(txtPosDescr2) = 0 Then
        MsgBox "Ceridian Key is a required field"
        txtPosDescr2.SetFocus
        Exit Function
    End If
End If

If Not glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    If Len(clpCode(1).Text) < 1 Then
        MsgBox "Status code is a required field"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Status code must be valid"
     clpCode(1).SetFocus
    Exit Function
End If

'Ticket #29069 Franks 08/15/2016 - begin
If txtJobCode.Text = "IND000" Then 'Ticket #30313 Franks 07/03/2017
Else
If Len(clpReportsTo(0).Text) < 1 Then
    MsgBox lblReptAuthor(0).Caption & " is a required field"
    clpReportsTo(0).SetFocus
    Exit Function
End If
End If
If clpReportsTo(0).Caption = "Unassigned" Then
    MsgBox lblReptAuthor(0).Caption & " must be valid"
    clpReportsTo(0).SetFocus
    Exit Function
End If
'Ticket #29069 Franks 08/15/2016 - end

'Ticket #28863 Franks 08/03/2016 - begin
If Len(dlpSDATE.Text) < 1 Then
    MsgBox "Start Date is a required field."
    dlpSDATE.SetFocus
    Exit Function
End If
If Not IsDate(dlpSDATE.Text) Then
    MsgBox "Invalid Start Date"
    dlpSDATE.SetFocus
    Exit Function
End If
If clpCode(1).Text = "INAC" Then
    If Len(dlpEDATE.Text) < 1 Then
        MsgBox "End Date is a required field if Position Status is Inactive."
        dlpEDATE.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpEDATE.Text) Then
        MsgBox "Invalid End Date"
        dlpEDATE.SetFocus
        Exit Function
    End If
Else
    If IsDate(dlpEDATE.Text) Then
        MsgBox "Cannot enter End Date if Position Status is not Inactive."
        dlpEDATE.Text = "" 'Ticket #29438 Franks 11/09/2016
        dlpEDATE.SetFocus
        Exit Function
    End If
End If
'Ticket #28863 Franks 08/03/2016 - end

If Not glbCompSerial = "S/N - 2415W" And Not glbCompSerial = "S/N - 2418W" And Not glbCompSerial = "S/N - 2433W" Then
'Ticket #16982 SPC- Volunteer System
'Ticket #17280 Charton-Hobbs Inc
'Ticket #21683 Kerry's Place Autism Services
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "Group code is a required field"
         clpCode(2).SetFocus
        Exit Function
    End If
End If

If clpCode(2).Caption = "Unassigned" Then
    MsgBox "Group code must be valid"
     clpCode(2).SetFocus
    Exit Function
End If
For X = 0 To 3 '2
    If Len(clpReportsTo(X).Text) > 0 Then
        If clpReportsTo(X).Caption = "Unassigned" Then
            MsgBox "Report to not a valid Position"
             clpReportsTo(X).SetFocus
            Exit Function
        End If
    End If
Next
If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox lblUnion.Caption & " code must be valid"
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If Len(clpNationalClass.Text) > 0 Then
    If clpNationalClass.Caption = "Unassigned" Then
        MsgBox "N.O.C. code must be valid"
        clpNationalClass.SetFocus
        Exit Function
    End If
End If

If glbLinamar Then
    If Len(clpCode(4).Text) > 0 Then
        If clpCode(4).Caption = "Unassigned" Then
            MsgBox "WF1 code must be valid"
             clpCode(4).SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(5).Text) > 0 Then
        If clpCode(5).Caption = "Unassigned" Then
            MsgBox "WF2 code must be valid"
             clpCode(5).SetFocus
            Exit Function
        End If
    End If
End If
' consider 0 same as empty entry
intLastNonZero = 0
If Len(txtNoPos) > 0 And Not IsNumeric(txtNoPos) Then
    If txtNoPos.Enabled Then
         MsgBox "Number of positions must be numeric."
         txtNoPos.SetFocus
         Exit Function
    End If
End If

If medFTENum.Enabled = True Then
    If Len(medFTENum) > 0 Then
        If Not IsNumeric(medFTENum) Then
             MsgBox "You must enter FTE Numeric"
             medFTENum.SetFocus
             Exit Function
        End If
    End If
End If

'Ticket #21378
'Ticket #21850
If medFTEHrs.Visible = True Then
    If medFTEHrs.Enabled = True Then
        If glbVadim And Left(comPayPer.Text, 1) = "A" Then
            If Not IsNumeric(medFTEHrs.Text) Then
                MsgBox "You must enter FTE Hours/Year to convert Annual Salary to Hourly for Vadim tranfer.", vbExclamation
                medFTEHrs.SetFocus
                Exit Function
            ElseIf Val(medFTEHrs.Text) = 0 Then
                MsgBox "FTE Hours/Year must be greater than 0 to convert Annual Salary to Hourly for Vadim tranfer.", vbExclamation
                medFTEHrs.SetFocus
                Exit Function
            End If
        End If
    
        If Len(medFTEHrs) > 0 Then
            If Not IsNumeric(medFTEHrs) Then
                MsgBox "You must enter FTE Hours/Year"
                medFTEHrs.SetFocus
                Exit Function
            End If
        End If
    End If
End If

' dkostka - 07/09/2001 - Added check for null Union field to next line, previously if union was null
'   step checks were skipped, as any comparison w/ a NULL is false.
If Not glbWFC Or (glbWFC And (clpCode(3).Text <> "NONE" And clpCode(3).Text <> "EXEC")) Or (glbWFC And clpCode(3) = "") Then 'Jaddy 10/21/99
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        If Len(medPayScale(X)) > 0 Then
            If Not IsNumeric(medPayScale(X)) Then
                MsgBox "Grid Step " & CStr(X) & " must be numeric"
                medPayScale(X).SetFocus
                Exit Function
            End If
            If CCur(medPayScale(X)) < 0 Then
                MsgBox "Grid Step " & CStr(X) & " can not be negative"
                medPayScale(X).SetFocus
                Exit Function
            End If
            If X >= 2 Then
                If CCur(medPayScale(X)) > 0 Then
                    If intLastNonZero > 0 And medPayScale(intLastNonZero + 1) <> 0 Then
                        If CCur(medPayScale(intLastNonZero)) > CCur(medPayScale(X)) Then
                            MsgBox "Grid Step must be entered in ascending order"
                            medPayScale(X).SetFocus
                            Exit Function
                        End If
                    End If
                    intLastNonZero = X
                End If
            Else
                If CCur(medPayScale(X)) > 0 Then
                    intLastNonZero = X
                End If
            End If
        Else
            medPayScale(X) = 0
        End If
    Next X
   
End If

'Ticket #25911 Franks 10/01/2014
If glbWFC Then
    For X = 7 To 9
        If Len(clpCode(X).Text) > 0 Then
            If clpCode(X).Caption = "Unassigned" Then
                MsgBox "code must be valid"
                clpCode(X).SetFocus
                Exit Function
            End If
        End If
    Next
    If Len(clpDiv.Text) > 0 Then
        If clpDiv.Caption = "Unassigned" Then
            MsgBox lStr("Division") & " code must be valid"
            clpDiv.SetFocus
            Exit Function
        End If
    End If
    If Len(txtMercerNo.Text) > 0 Then
        'If Not IsNumeric(txtMercerNo.Text) Then
        '    MsgBox "Mercer Survey # code must be numeric"
        '    txtMercerNo.SetFocus
        '    Exit Function
        'End If
    End If
    If Len(txtJobCode.Text) > 0 Then
        If lblJobCodeDesc.Caption = "Unassigned" Then
            MsgBox "Job Code must be valid"
            txtJobCode.SetFocus
            Exit Function
        End If
    End If

    If Left(txtPosDescr.Text, 2) = "Z " Then
        If Not clpCode(1).Text = "INAC" Then
            clpCode(1).Text = "INAC"
        End If
    End If
End If

chkPositions = True

Exit Function

chkPositions_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkPositions", "HRJOB", "edit/Add")
Resume Next
Call RollBack '14June99 js

End Function

Private Sub chkHideInactive_Click()
    Dim SQLQ  As String
    
    SQLQ = "SELECT * FROM HRJOB WHERE 1=1 "
    If Len(glbWFCUserSecList) > 0 Then 'Ticket #27609 Franks 10/13/2015
        SQLQ = SQLQ & " AND JB_SECTION IN " & glbWFCUserSecList & " "
    End If
    If chkHideInactive Then
        SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
        If glbOracle Then 'Ticket #16416
            SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
        ElseIf glbSQL Then
            SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
        Else
            SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
        End If
    End If

    If glbOracle Then
        SQLQ = SQLQ & " ORDER BY UPPER(JB_DESCR)"
    Else
        SQLQ = SQLQ & " ORDER BY JB_DESCR"
    End If

    Data1.RecordSource = SQLQ
    Data1.Refresh
'    Set fRS = Data1.Recordset.Clone
'    vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If glbWFC And fglbNew Then 'Ticket #29846 Franks 03/06/2017
        If Index = 6 Or Index = 8 Then 'Band or MarketLine
            If Len(clpCode(6).Text) > 0 Then
                MskApprLimit.Text = WFCSigningApprovalLimitGet(clpCode(7).Text, clpCode(6).Text, clpCode(8).Text)
            End If
        End If
    End If
End Sub

Private Sub clpDiv_LostFocus()
    If glbWFC And fglbNew Then
        If Len(clpDiv.Text) > 0 Then
            'Call getDataFromJobMaster(clpDiv.Text)
            Call getDataFromDivMaster(clpDiv.Text)
        End If
    End If
End Sub

Private Sub clpGrid_LostFocus()
    If txtPosition <> "" And clpGrid <> "" Then
        txtLambtonJob = Left(clpGrid, 1) & txtPosition & Mid(clpGrid, 2, 1)
    End If
End Sub

Private Sub clpReportsTo_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmbBand_Click()
'txtBand = cmbBand       'Jaddy 8/999
'End Sub

'Private Sub cmbBand_GotFocus()
'Call SetPanHelp(ActiveControl) 'Jaddy 8/999
'End Sub

Private Sub cmbMidPoint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbMidPoint_LostFocus()
txtMidPoint = cmbMidPoint.ListIndex + 1
End Sub

Public Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

Exit Sub

Can_Err:
If Err = 3058 Then
Err = 0
Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
Resume Next
Call RollBack '14June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdClose_Click()

glbDept = ""
glbDeptDesc = ""
Unload Me

End Sub

Private Sub cmdAttachJobFiles_Click()
    glbDocName = "JobFiles"
    frmJobDocument.Show 1
    DoEvents
End Sub

Private Sub cmdCopy2AnotheDiv_Click()
    glbWFC_IPPopFormName = "WFCPosMasterAndBudCopByDiv"
    frmBENGRCopy.Show 1
End Sub

Private Sub cmdCopy2AnotherPlant_Click() 'Ticket #29846 Franks 03/07/2017
    glbWFC_IPPopFormName = "WFCPosMasterAndBudCopy"
    frmBENGRCopy.Show 1
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdCountPos_Click()

On Error GoTo CountErr

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If mod_Upd_Pos_Totals(True) Then
        Beep
        MsgBox "Positions Counted"
    End If
    Data1.Refresh
    modChekCount
End If

Exit Sub

CountErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Count Pos Error", "HRJOB Refresh", "Refresh")
Resume Next
Call RollBack '14June99 js

End Sub

Public Sub cmdDelete_Click()
Dim X, Msg$, Job, SQLQ As String, ReptsTo$
Dim snapAssJobs As New ADODB.Recordset
Dim snapAssJobsTerm As New ADODB.Recordset
Dim snapAssGrid As New ADODB.Recordset
On Error GoTo DelErr

Job = Data1.Recordset("JB_CODE")
Screen.MousePointer = HOURGLASS
ReptsTo$ = modReportsTo(Job)
Screen.MousePointer = DEFAULT

If Not ReptsTo$ = "None" Then
    Msg$ = "Position  " & Chr(10) & ReptsTo$
    Msg$ = Msg$ & Chr(10) & "still reports to this position!"
    MsgBox Msg$
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If glbOracle Then
    SQLQ = "SELECT HR_JOB_HISTORY.JH_JOB, ED_SURNAME || ', ' || ED_FNAME AS Name, HREMP.ED_EMPNBR, HR_JOB_HISTORY.JH_DHRS FROM HR_JOB_HISTORY,HREMP WHERE HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR(+) "
    SQLQ = SQLQ & " AND JH_JOB = '" & Job & "'"
Else
    SQLQ = "SELECT HR_JOB_HISTORY.JH_JOB, ED_SURNAME + ', ' + ED_FNAME AS Name, HREMP.ED_EMPNBR, HR_JOB_HISTORY.JH_DHRS FROM HR_JOB_HISTORY INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " WHERE JH_JOB = '" & Job & "'"
End If

snapAssJobs.Open SQLQ, gdbAdoIhr001, adOpenStatic
Screen.MousePointer = DEFAULT
'Ticket# 6685 Check Terminated employees as well.
If glbOracle Then
    SQLQ = "SELECT Term_JOB_HISTORY.JH_JOB, Term_HREMP.ED_SURNAME || ', ' || Term_HREMP.ED_FNAME AS Name, Term_HREMP.ED_EMPNBR, Term_JOB_HISTORY.JH_DHRS FROM Term_JOB_HISTORY,Term_HREMP WHERE Term_JOB_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ(+) "
    SQLQ = SQLQ & " AND Term_JOB_HISTORY.JH_JOB = '" & Job & "'"
Else
    SQLQ = "SELECT Term_JOB_HISTORY.JH_JOB, Term_HREMP.ED_SURNAME + ', ' + Term_HREMP.ED_FNAME AS Name, Term_HREMP.ED_EMPNBR, Term_JOB_HISTORY.JH_DHRS FROM Term_JOB_HISTORY INNER JOIN Term_HREMP ON Term_JOB_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " WHERE Term_JOB_HISTORY.JH_JOB = '" & Job & "'"
End If
snapAssJobsTerm.Open SQLQ, gdbAdoIhr001X, adOpenStatic

SQLQ = "SELECT JB_GRID,JB_ID FROM HRJOB_GRADE WHERE JB_CODE='" & Job & "'"
snapAssGrid.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapAssJobs.BOF And snapAssJobs.EOF And snapAssJobsTerm.BOF And snapAssJobsTerm.EOF And snapAssGrid.EOF Then
    Msg$ = "Are You Sure You Want To Delete This Record?"
    Msg$ = Msg$ & Chr(10) & "All of its Skills and Evaluation Factors "
    Msg$ = Msg$ & Chr(10) & "will also be deleted!"
    X = MsgBox(Msg, 36, "Confirm Delete")
    If X <> 6 Then Exit Sub
    
    If Not (glbWFC And glbPlantCode = "GREN") Then   'Greensboro
        Call Codes_Master_Integration("POSITION", txtPosition, , True)
    End If
    
    'Delete from Vadim as well
    If glbVadim Then
        'Delete the Occupation Rates first - Step Amounts
        If Not glbMultiGrid Then
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'For X = 1 To 11
            'For X = 1 To 15
            For X = 1 To 20
                DoEvents
                
                If Val(medPayScale(X).Text) > 0 Then
                    Call Passing_Salary_Grid_Vadim(X, medPayScale(X).Text, 0, Date, txtPosition)
                End If
            Next X
        End If
        
        'Delete from Occupation table
        Call Passing_Position_Master_Vadim(txtPosition, "D", "", "")
    End If
    
    gdbAdoIhr001.BeginTrans

    SQLQ = "DELETE FROM HRJOBEVL WHERE JE_CODE = '" & txtPosition & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "DELETE FROM HRJOBSKL WHERE JS_CODE = '" & txtPosition & "' "
    gdbAdoIhr001.Execute SQLQ
    'If Not glbSQL And Not glbOracle Then
        SQLQ = "DELETE FROM HR_JOB_COURSE WHERE PC_JOB = '" & txtPosition & "' "
        gdbAdoIhr001.Execute SQLQ
    'End If
    
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    
    'George Feb 3,2006 #10266
    If gsAttachment_DB Then
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute "Delete from HRDOC_JOB where DB_TYPE='" & UCase(glbDocName) & "' and DB_JOB='" & glbPos & "'"
    gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Feb 3,2006 #10266

    
    Data1.Refresh
    
Else
    Msg$ = ""
    'For active employees
    If Not snapAssJobs.EOF Then
        Msg$ = Msg$ & "This job is presently assigned to active employees: "
    End If
    X = 0
    While Not snapAssJobs.EOF And X < 10
        Msg$ = Msg$ & Chr(10) & snapAssJobs("Name") & " -  # " & snapAssJobs("ED_EMPNBR")
        X = X + 1
        snapAssJobs.MoveNext
    Wend
    'For Terminated employees
    If Not snapAssJobsTerm.EOF Then
        Msg$ = Msg$ & Chr(10) & "This job is presently assigned to terminated employees: "
    End If
    X = 0
    While Not snapAssJobsTerm.EOF And X < 10
        Msg$ = Msg$ & Chr(10) & snapAssJobsTerm("Name") & " -  # " & snapAssJobsTerm("ED_EMPNBR")
        X = X + 1
        snapAssJobsTerm.MoveNext
    Wend
    
    'For Salary Grid
    If glbMultiGrid Then
        If Not snapAssGrid.EOF Then
            Msg$ = Msg$ & Chr(10) & "This job is presently assigned to Salary Grids: "
        End If
        X = 0
        While Not snapAssGrid.EOF And X < 10
            Msg$ = Msg$ & Chr(10) & snapAssGrid("JB_GRID")
            X = X + 1
            snapAssGrid.MoveNext
        Wend
    End If
    Msg$ = Msg$ & Chr(10) & "Record will not be deleted."
    MsgBox Msg$
        
    Exit Sub
End If
'Call modSTUPD(True)
fglbNew = False
Call SET_UP_MODE


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRPROV", "Delete")
Resume Next
Call RollBack '14June99 js

End Sub



Public Sub cmdNew_Click()
Dim X As Integer
Dim xNewPosCode

On Error GoTo NewErr
'Call modSTUPD(False)

'Ticket #27827 Franks 12/02/2015 - begin
glbWFCNewPosJob = ""
glbWFCNewPosDiv = ""
glbWFCNewPosStatus = ""
frmNewPosWFC.Show 1
If glbWFCNewPosJob = "" Then
    Exit Sub
Else
    'get the new position code
    xNewPosCode = getNewPosCode(glbWFCNewPosDiv, glbWFCNewPosStatus)
End If
'Ticket #27827 Franks 12/02/2015 - end

fglbNew = True
Call SET_UP_MODE

'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
End If
'George on Jan 26,2006 #10266


Call Set_Control("B", Me)
rsDATA.AddNew

txtComp.Text = "001"
lblSalCode = "A"
comPayPer.ListIndex = 0
lblPoints = 0
lblPosFiled = 0
oJobCode = ""
oJobDesc = ""
oJobUnion = ""
oStatus = ""
clpCode(6).Text = ""
If glbMultiGrid Then
    clpGrid.Visible = True
    lblGridC.Visible = True
    clpGrid.Text = ""
    If glbLambton Then
        txtLambtonJob.Visible = True
        lblLambtonJob.Visible = True
        txtLambtonJob = ""
    End If
End If
'txtPosition.Enabled = True
'txtPosition.SetFocus
Call INI_GridStep

'Ticket #27827 Franks 12/02/2015 - begin
txtPosition.Text = xNewPosCode
txtJobCode.Text = glbWFCNewPosJob
If Not Left(glbWFCNewPosJob, 1) = "U" Then 'Ticket #29955 Franks 03/20/2017
    '"   New Position should have the Band updated. The Band is the 3rd position of the Job Code. If the Job Code begins with U, no band
    clpCode(6).Text = Mid(glbWFCNewPosJob, 3, 1)
End If

Call txtJobCode_LostFocus
clpDiv.Text = glbWFCNewPosDiv
Call clpDiv_LostFocus
txtPosDescr.Text = lblJobCodeDesc.Caption
txtPosDescr2.Text = lblJobCodeDesc.Caption
'Ticket #27827 Franks 12/02/2015 - end

'Ticket #28118 Franks 02/01/2016
medFTENum.Text = 1
medFTEHrs.Text = 2028
medHours.Text = 8


Exit Sub

NewErr:
If Err = 3058 Then
Err = 0
Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRJOB", "AddNew")
Resume Next
Call RollBack '14June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()
Dim strLastPos, SQLQ, X, I
Dim xlocNew As Boolean
Dim Msg$, xOldPosCode, xOldPosEDate

On Error GoTo OK_Err

'''Call WFCOldPosUpt("", "1050PCS02078", Date) '??? TEST

xlocNew = fglbNew

'Ticket #29438 Franks 11/08/2016
'"   Made a position from inactive to active….position above. There was no Budgeted Position Master for this record. The system should automatically go to Budgeted Position like it does when creating a new position -
If oStatus = "INAC" Then
    If Not oStatus = clpCode(1).Text Then
        xlocNew = True 'treat it as new so the program goes to Budgested Position automatically
    End If
End If

For I = medPayScale.LBound To medPayScale.UBound
    If medPayScale(I).Text = "" Then medPayScale(I).Text = "0.00"
Next I

If Not chkPositions() Then Exit Sub
If glbVadim Then
    Call Transfer_Position_Master_Vadim
End If
If Not Data1.Recordset.EOF Then
    lstMidPoint = Data1.Recordset("JB_MIDPOINT")    'Hemu
End If

strLastPos = txtPosition

If xlocNew Then
    glbPos = txtPosition.Text
    glbPosDesc = txtPosDescr.Text
    glbJobSection = clpCode(7).Text
End If

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbLinamar Then rsDATA("JB_WF_TABL") = "LNWF"
rsDATA("JB_MIDPOINT") = Val(txtMidPoint)
If rsDATA("JB_MIDPOINT") = 0 Then rsDATA("JB_MIDPOINT") = 1

If glbLinamar Then rsDATA!JB_WF_TABL = "LNWF"

rsDATA!JB_MIDPOINT = Val(txtMidPoint)
If rsDATA!JB_MIDPOINT = 0 Then rsDATA!JB_MIDPOINT = 1

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

SQLQ = "JB_CODE = '" & strLastPos & "'"
Data1.Recordset.Find SQLQ

If glbWFC Then 'Ticket #29183 Franks 09/12/2016
    If fglbNew Then
        Msg$ = "Is this position a replacement for an existing position? "
        X = MsgBox(Msg, 36, "Confirm")
        If X = 6 Then  'Yes
            glbChgTermDate = ""
            glbChgTermReason = strLastPos
            frmMsgTerm.PenTermDate = "WFCPosEnd_Change"
            frmMsgTerm.Show 1
            If IsDate(glbChgTermDate) Then
                xOldPosCode = glbChgTermReason
                xOldPosEDate = glbChgTermDate
                'Call WFCOldPosUpt(strLastPos, glbChgTermReason, glbChgTermDate)
                
                gdbAdoIhr001W.BeginTrans
                gdbAdoIhr001W.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
                gdbAdoIhr001W.CommitTrans
        
                glbSPCPPay = strLastPos
                glbWFC_IPPopFormName = "WFCEmpListWithOldPos"
                glbWFC_IncePlanID = 0
                frmCheckListView.Show 1
                'If glbWFC_IncePlanID = -1 Then
                '    'Cancel, do nothing
                'End If
                
                Call WFCReptChaUptPopUp 'Ticket #29484 Franks 11/22/2016

                Call WFCOldPosUpt(strLastPos, xOldPosCode, xOldPosEDate)
            End If
        End If
    End If
End If

'Call modSTUPD(True)
Call modChekCount
EmpChgErrors = ""
If glbMultiGrid Then
    If fglbNew Then Call ADDGRIDS(strLastPos)
Else
    SkipResetGridStep = True
    DoEvents
    Call GridStepChange
    DoEvents
    SkipResetGridStep = False
    ' dkostka - 01/29/2002 - Added list of employees that couldn't be changed for grid step changes.
    If EmpChgErrors <> "" Then
        MsgBox "The following employee salaries could not be changed due to missing 'Hours Per Week' values:" & vbCrLf & EmpChgErrors
    End If
End If

If glbWFC And fglbNew Then 'Ticket #27827 Franks 12/02/2015
    Call WFCNextPosNoSetup("Ongoing")
End If

fglbNew = False

If Not (glbWFC And glbPlantCode = "GREN") Then   'Greensboro
    Call Codes_Master_Integration("POSITION", txtPosition)
End If



Call SET_UP_MODE
Call Display_Value

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            'glbDocKey = xID
            glbPos = rsDATA("JB_CODE") 'Data1.Recordset("JB_CODE")
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
            Call DispimgIcon(Me, "frmECounsel")
        End If
    End If
    glbDocImpFile = ""
End If

If glbWFC And xlocNew Then 'Ticket #28340 Franks 03/21/2016
    Load frmPosBudgetWFC
    Unload frmMPOSITIONSWFC
End If

Exit Sub

OK_Err:
If Err = 3022 Then  'trying to add a duplicate key
    Err = 0
    MsgBox "Position already exists"
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOB", "Update")
Resume Next
Call RollBack '14June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub ADDGRIDS(zJOB)
Dim rsGrid As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim X
rsGrid.Open "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & zJOB & "'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If clpGrid = "DFLT" Then
    rsTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='JBGD' AND TB_KEY='DFLT'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsTABL.EOF Then
        rsTABL.AddNew
        rsTABL("TB_NAME") = "JBGD"
        rsTABL("TB_KEY") = "DFLT"
        rsTABL("TB_DESC") = "DEFAULT"
        rsTABL.Update
    End If
    rsTABL.Close
End If
If rsGrid.EOF Then
    rsGrid.AddNew
    rsGrid("JB_CODE") = zJOB
    rsGrid("JB_GRID") = clpGrid
    rsGrid("JB_MIDPOINT") = "1"
    rsGrid("JB_SALCD") = "A"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        rsGrid("JB_S" & X) = 0
    Next
    rsGrid.Update
End If
rsGrid.Close
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Positions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Positions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub combMidPoint()

Dim I, xLev, X
X = cmbMidPoint.ListIndex
cmbMidPoint.Clear
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For I = 1 To 11
'For I = 1 To 15
For I = 1 To 20
    If glbCompSerial = "S/N - 2191W" Then
        If Val(txtPosDescr2) = 0.5 Then
            xLev = Format((I + 1) / 2, "0.0")
        Else
            xLev = Format(I, "0.0")
        End If
    Else
        xLev = CStr(I)
    End If
    cmbMidPoint.AddItem xLev
    lblGrid(I - 1) = xLev
Next I
If X <= 0 Then X = 0
cmbMidPoint.ListIndex = X  'Display first item in list

If glbCompSerial = "S/N - 2366W" Then   'Family Youth and Child Services of Muskoka
    lblGrid(0).Caption = "Start"
    lblGrid(1).Caption = "1"
    lblGrid(2).Caption = "2"
    lblGrid(3).Caption = "3"
    lblGrid(4).Caption = "4"
    lblGrid(5).Caption = "5"
    lblGrid(6).Caption = "6"
    lblGrid(7).Caption = "7"
    lblGrid(8).Caption = "8"
    lblGrid(9).Caption = "9"
    lblGrid(10).Caption = "10"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    lblGrid(11).Caption = "11"
    lblGrid(12).Caption = "12"
    lblGrid(13).Caption = "13"
    lblGrid(14).Caption = "14"
    
    lblGrid(15).Caption = "15"
    lblGrid(16).Caption = "16"
    lblGrid(17).Caption = "17"
    lblGrid(18).Caption = "18"
    lblGrid(19).Caption = "19"
    
End If

End Sub


Private Sub cmdExp_Click()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rs As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim RsHRPARCO As New ADODB.Recordset
Dim ImportFile, xlsFileTmp
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String
Dim xNum As Integer
Dim xRows As Long
Dim xRow As Long
Dim xEmpnbr
Dim xFlag As Boolean
Dim xYear, xPlant, xPos As String, xDiv, xBUnit, xDept As String, xGL As String, xBudPos, xFte, xFTEHrs, xUptMsg
Dim xTmp, I

    If Not gSec_Upd_Position Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    'Budgeted Position Import.xls or Budgeted Position Import.xlsx
    'check file name
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Master Export Tmp.xls"
    ImportFile = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Master Export.xls"
    If Dir(xlsFileTmp) = "" Then
      MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
      Exit Sub
    End If
    
    Msg = "This program will export Position Master records into this file: "
    Msg = Msg & Chr(10) & ImportFile
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do it? "
    a% = MsgBox(Msg, 36, "Confirm")
    
    If a% <> 6 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    If (Dir(ImportFile)) <> "" Then Kill ImportFile

    FileCopy xlsFileTmp, ImportFile
    
    SQLQ = "SELECT * FROM HRJOB WHERE (1=1) "
    If Len(clpCode(12).Text) > 0 Then 'Ticket #29552 Franks 12/14/2016
        SQLQ = SQLQ & "AND JB_SECTION = '" & clpCode(12).Text & "' "
    End If
    SQLQ = SQLQ & "ORDER BY JB_SECTION, JB_DESCR "
    
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        MsgBox "There is no record in HRJOB table."
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
        
    xRow = 2
    xNum = 0
    xRows = rs.RecordCount
    Do While Not rs.EOF
        MDIMain.panHelp(0).FloodPercent = (xNum / xRows) * 100
        DoEvents
        exSheet.Cells(xRow, 1) = rs("JB_CODE")
        exSheet.Cells(xRow, 2) = rs("JB_DESCR")
        If Not IsNull(rs("JB_DESCR2")) Then exSheet.Cells(xRow, 3) = rs("JB_DESCR2")
        exSheet.Cells(xRow, 4) = rs("JB_DIV")
        If Not IsNull(rs("JB_REGION")) Then exSheet.Cells(xRow, 5) = rs("JB_REGION")
        If Not IsNull(rs("JB_SECTION")) Then exSheet.Cells(xRow, 6) = rs("JB_SECTION")
        
        If Not IsNull(rs("JB_MERCER_NO")) Then exSheet.Cells(xRow, 7) = rs("JB_MERCER_NO")
        If Not IsNull(rs("JB_JOBCODE")) Then exSheet.Cells(xRow, 8) = rs("JB_JOBCODE")
        If Not IsNull(rs("JB_POINTS")) Then exSheet.Cells(xRow, 9) = rs("JB_POINTS")
        If Not IsNull(rs("JB_STATUS")) Then exSheet.Cells(xRow, 10) = rs("JB_STATUS")
        If Not IsNull(rs("JB_GRPCD")) Then exSheet.Cells(xRow, 11) = rs("JB_GRPCD")
        If Not IsNull(rs("JB_LEVEL")) Then exSheet.Cells(xRow, 12) = rs("JB_LEVEL")
        If Not IsNull(rs("JB_BAND")) Then exSheet.Cells(xRow, 13) = rs("JB_BAND")
        If Not IsNull(rs("JB_ORG")) Then exSheet.Cells(xRow, 14) = rs("JB_ORG")
        If Not IsNull(rs("JB_MARKETLINE")) Then exSheet.Cells(xRow, 15) = rs("JB_MARKETLINE")
        If Not IsNull(rs("JB_FEDGRP")) Then exSheet.Cells(xRow, 16) = rs("JB_FEDGRP")
        If Not IsNull(rs("JB_REPTAU")) Then exSheet.Cells(xRow, 17) = rs("JB_REPTAU")
        If Not IsNull(rs("JB_REPTAU2")) Then exSheet.Cells(xRow, 18) = rs("JB_REPTAU2")
        If Not IsNull(rs("JB_REPTAU3")) Then exSheet.Cells(xRow, 19) = rs("JB_REPTAU3")
        If Not IsNull(rs("JB_REPTAU4")) Then exSheet.Cells(xRow, 20) = rs("JB_REPTAU4")
        If Not IsNull(rs("JB_FTENUM")) Then exSheet.Cells(xRow, 21) = rs("JB_FTENUM")
        If Not IsNull(rs("JB_FTETOTNU")) Then exSheet.Cells(xRow, 22) = rs("JB_FTETOTNU")
        If Not IsNull(rs("JB_FTEHRS")) Then exSheet.Cells(xRow, 23) = rs("JB_FTEHRS")
        If Not IsNull(rs("JB_FTETOTHR")) Then exSheet.Cells(xRow, 24) = rs("JB_FTETOTHR")
        If Not IsNull(rs("JB_DHRS")) Then exSheet.Cells(xRow, 25) = rs("JB_DHRS")
        'Ticket #29484 Franks 11/21/2016 - begin
        If Not IsNull(rs("JB_SDATE")) Then exSheet.Cells(xRow, 26) = rs("JB_SDATE")
        If Not IsNull(rs("JB_EDATE")) Then exSheet.Cells(xRow, 27) = rs("JB_EDATE")
        If Not IsNull(rs("JB_USERDEF1")) Then exSheet.Cells(xRow, 28) = rs("JB_USERDEF1")
        If Not IsNull(rs("JB_USERDEF2")) Then exSheet.Cells(xRow, 29) = rs("JB_USERDEF2")
        If Not IsNull(rs("JB_POSTYPE")) Then exSheet.Cells(xRow, 30) = rs("JB_POSTYPE")
        If Not IsNull(rs("JB_LDATE")) Then exSheet.Cells(xRow, 31) = rs("JB_LDATE")
        If Not IsNull(rs("JB_LUSER")) Then exSheet.Cells(xRow, 32) = rs("JB_LUSER")
        'Ticket #29484 Franks 11/21/2016 - end
        xNum = xNum + 1
        xRow = xRow + 1
        rs.MoveNext
    Loop
    rs.Close
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""


    Call Pause(1)
    
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = ""
    
    Screen.MousePointer = vbDefault

    I = 0
    SQLQ = "SELECT * FROM HRPARCO "
    RsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not RsHRPARCO.EOF Then
        I = RsHRPARCO("PC_NEXT_POS_NBR")
    End If
    RsHRPARCO.Close
        
    xTmp = "Finished!   " & Chr(10) & Chr(10) & "Next Available Number is " & I
    MsgBox xTmp ' "   Finished!   "

End Sub

Private Sub cmdImp_Click()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rs As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim ImportFile
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String
Dim xNum As Integer
Dim xRows As Long
Dim xRow As Long
Dim xEmpnbr
Dim xFlag As Boolean
Dim xYear, xPlant, xPos As String, xDiv, xBUnit, xDept As String, xGL As String, xBudPos#, xFte, xFTEHrs, xUptMsg, xAddOrUpt
Dim xTmp

    If Not gSec_Upd_Position Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    'Budgeted Position Import.xls or Budgeted Position Import.xlsx
    'check file name
    ImportFile = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC Position Master Import.xls"
    If Dir(ImportFile) = "" Then
      MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
      Exit Sub
    End If
    
    Msg = "This program will load Positions from the file: "
    Msg = Msg & Chr(10) & ImportFile
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do it? "
    a% = MsgBox(Msg, 36, "Confirm")
    
    If a% <> 6 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    MDIMain.panHelp(0).FloodType = 1
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
    xRows = getRows(exSheet)
    
    xFlag = True
    xPos = exSheet.Cells(1, 4) 'exSheet.Cells(2, 31)
    If Not Trim(xPos) = "Division" Then
        'MsgBox ("Year is a required field")
        xFlag = False
    End If
    If xFlag = False Then
        MsgBox "Invalid File"
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Exit Sub
    End If
    
    
    For xRow = 2 To xRows
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xUptMsg = ""
        
        xPos = exSheet.Cells(xRow, 1)
        xFlag = True
        
        If Len(xPos) < 11 Or Len(xPos) > 12 Then
            xUptMsg = "Invalid Position Code length": GoTo Error_Upt
        End If
        
        xDiv = exSheet.Cells(xRow, 4)
        If Len(xDiv) = 0 Then
            xUptMsg = "No Division": GoTo Error_Upt
        Else
            xTmp = getDivDescPub(xDiv)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid Division Code": GoTo Error_Upt
            End If
        End If
        
        xBUnit = exSheet.Cells(xRow, 5)
        If Len(xBUnit) > 0 Then
            xTmp = GetTABLDesc("EDRG", xBUnit)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid Business Unit Code": GoTo Error_Upt
            End If
        End If
        
        xPlant = exSheet.Cells(xRow, 6)
        If Len(xPlant) = 0 Then
            xUptMsg = "Invalid Plant": GoTo Error_Upt
        Else
            xTmp = GetTABLDesc("EDSE", xPlant)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid Plant Code": GoTo Error_Upt
            End If
        End If
        
        'If xPos = "1001HR03839" Then
        'Debug.Print ""
        'End If

Error_Upt:
        If Len(xUptMsg) > 0 Then
            exSheet.Cells(xRow, 36) = xUptMsg
            GoTo Next_Rec
        End If
        
        xAddOrUpt = ""
        'Position/Division must be setup in HRJOB. If not, display a message saying "This Position has not been assigned to this Division". Dead stop.
        xPos = Trim(xPos)
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & Trim(xPos) & "' "
        'SQLQ = SQLQ & "AND JB_DIV = '" & xDiv & "' "
        If rsJOB.State <> 0 Then rsJOB.Close
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsJOB.EOF Then
            xAddOrUpt = "New"
            rsJOB.AddNew
            rsJOB("JB_CODE") = Left(xPos, 25)
        Else
            xAddOrUpt = "Update"
        End If
        xTmp = exSheet.Cells(xRow, 2)
        If Len(xTmp) > 0 Then rsJOB("JB_DESCR") = Left(xTmp, 100)
        xTmp = exSheet.Cells(xRow, 3)
        If Len(xTmp) > 0 Then rsJOB("JB_DESCR2") = Left(xTmp, 100)
        
        xTmp = exSheet.Cells(xRow, 4)
        If Len(xTmp) > 0 Then rsJOB("JB_DIV") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 5)
        If Len(xTmp) > 0 Then rsJOB("JB_REGION") = Left(xTmp, 20)
        xTmp = exSheet.Cells(xRow, 6)
        If Len(xTmp) > 0 Then rsJOB("JB_SECTION") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 7)
        If Len(xTmp) > 0 Then rsJOB("JB_MERCER_NO") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 8)
        If Len(xTmp) > 0 Then rsJOB("JB_JOBCODE") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 9)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_POINTS") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 10)
        If Len(xTmp) > 0 Then rsJOB("JB_STATUS") = Left(xTmp, 4)
        
        xTmp = exSheet.Cells(xRow, 11)
        If Len(xTmp) > 0 Then rsJOB("JB_GRPCD") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 12)
        If Len(xTmp) > 0 Then rsJOB("JB_LEVEL") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 13)
        If Len(xTmp) > 0 Then rsJOB("JB_BAND") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 14)
        If Len(xTmp) > 0 Then rsJOB("JB_ORG") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 15)
        If Len(xTmp) > 0 Then rsJOB("JB_MARKETLINE") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 16)
        If Len(xTmp) > 0 Then rsJOB("JB_FEDGRP") = Left(xTmp, 4)
        xTmp = exSheet.Cells(xRow, 17)
        If Len(xTmp) > 0 Then rsJOB("JB_REPTAU") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 18)
        If Len(xTmp) > 0 Then rsJOB("JB_REPTAU2") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 19)
        If Len(xTmp) > 0 Then rsJOB("JB_REPTAU3") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 20)
        If Len(xTmp) > 0 Then rsJOB("JB_REPTAU4") = Left(xTmp, 25)

        xTmp = exSheet.Cells(xRow, 21)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_FTENUM") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 22)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_FTETOTNU") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 23)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_FTEHRS") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 24)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_FTETOTHR") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 25)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_DHRS") = xTmp
        End If
        rsJOB("JB_LDATE") = Date
        rsJOB("JB_LTIME") = Time$
        rsJOB("JB_LUSER") = "PosMstImp"
        'Ticket #29484 Franks 11/21/2016 - begin
        xTmp = exSheet.Cells(xRow, 26)
        If Len(xTmp) > 0 Then
            If IsDate(xTmp) Then rsJOB("JB_SDATE") = CVDate(xTmp)
        End If
        xTmp = exSheet.Cells(xRow, 27)
        If Len(xTmp) > 0 Then
            If IsDate(xTmp) Then rsJOB("JB_EDATE") = CVDate(xTmp)
        End If
        xTmp = exSheet.Cells(xRow, 28)
        If Len(xTmp) > 0 Then rsJOB("JB_USERDEF1") = Left(xTmp, 25)
        xTmp = exSheet.Cells(xRow, 29)
        If Len(xTmp) > 0 Then
            If IsNumeric(xTmp) Then rsJOB("JB_USERDEF2") = xTmp
        End If
        xTmp = exSheet.Cells(xRow, 30)
        If Len(xTmp) > 0 Then rsJOB("JB_POSTYPE") = Left(xTmp, 20)
        
        'Ticket #29484 Franks 11/21/2016 - end
        'exSheet.Cells(xRow, 26) = xAddOrUpt
        exSheet.Cells(xRow, 36) = xAddOrUpt
        rsJOB.Update
        
        'Ticket #29484 Franks 11/21/2016
        'On a new position, auto create the budgeted position master with
        'Budgeted #/Pos'ns and FTE number will come from the last column of the spreadsheet. If no value is in the spreadsheet, default to 1.
        'FTE Hours/Year is always defaulted to 2080
        If xAddOrUpt = "New" Then
            xTmp = exSheet.Cells(xRow, 33) 'Budgeted #Pos'ns(New record only)
            xBudPos# = 0
            If Len(xTmp) > 0 Then
                If IsNumeric(xTmp) Then
                    xBudPos# = xTmp
                End If
            End If
            
            Call mod_ADD_Pos_Budget_WFC(xPos, xPlant, xDiv, xBudPos#, "PosMstImp")  'Ticket #29484 Franks 11/21/2016
            
        End If
Next_Rec:

    Next
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Finished!   " & Chr(10) & "Check column AJ to see if there are some errors."
    Unload Me

End Sub



Private Sub cmdUptSignApprov_Click() 'Ticket #29846 Franks 03/07/2017
    glbWFC_IPPopFormName = "WFCUptSigningApproval"
    frmBENGRCopy.Show 1
End Sub

Private Sub comPayPer_Click()
Select Case comPayPer.ListIndex
Case 0: lblSalCode.Caption = "A"
        fraGrid.Caption = "Annual Grid"
        fraGrid1.Caption = "Hourly Grid"
Case 1: lblSalCode.Caption = "H"
        fraGrid.Caption = "Hourly Grid"
        fraGrid1.Caption = "Annual Grid"
Case Else: lblSalCode.Caption = ""
        fraGrid.Caption = "Grid Steps"
        fraGrid1.Caption = "Grid Steps"
End Select
'Hemu - Ticket #10139 - Town of Aurora only
'Oxford Ticket #15590
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2259W" Then
    Call Calculate_Secondary_Grid_Steps(0)
End If

End Sub

Private Sub comPayPer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayPer_KeyPress(KeyAscii As Integer)

If comPayPer.Text = "Annual Grid" Then
    lblSalCode.Caption = "A"
    fraGrid.Caption = "Annual Grid"
    fraGrid1.Caption = "Hourly Grid"
ElseIf comPayPer.Text = "Hourly Grid" Then
    lblSalCode.Caption = "H"
    fraGrid.Caption = "Hourly Grid"
    fraGrid1.Caption = "Annual Grid"
End If

End Sub

Private Sub comPayPer_LostFocus()

If comPayPer.ListIndex = 0 Then
    lblSalCode.Caption = "A"
    fraGrid.Caption = "Annual Grid"
    fraGrid1.Caption = "Hourly Grid"
ElseIf comPayPer.ListIndex = 1 Then
    lblSalCode.Caption = "H"
    fraGrid.Caption = "Hourly Grid"
    fraGrid1.Caption = "Annual Grid"
End If

End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    glbOnTop = "frmMPOSITIONSWFC"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmMPOSITIONSWFC"
End Sub

Private Sub Form_Load()
flagLoad = 1
Dim SQLQ As String, RFound%, X As Integer

If Not glbWFC Then 'Ticket #20479 Franks 07/11/2011, this field is for WFC only
    vbxTrueGrid.Columns(5).Visible = False
End If

cmdImp.Enabled = gSec_Upd_Job_Master 'Ticket #29438 Franks 12/01/2016
cmdExp.Enabled = gSec_Upd_Job_Master 'Ticket #29438 Franks 12/01/2016
cmdCopy2AnotherPlant.Enabled = gSec_Upd_Job_Master 'Ticket #29438 Franks 12/01/2016
cmdCopy2AnotheDiv.Enabled = gSec_Upd_Job_Master 'Ticket #29846 Franks 03/06/2017
cmdUptSignApprov.Enabled = gSec_Upd_Job_Master 'Ticket #29846 Franks 03/06/2017

glbOnTop = "frmMPOSITIONSWFC"
Data1.ConnectionString = glbAdoIHRDB

'Hemu - Ticket #10139 - Town of Aurora only
'Frank- Ticket #15590 - County of Oxford
If glbCompSerial <> "S/N - 2378W" And glbCompSerial <> "S/N - 2259W" Then
    fraGrid1.Visible = False
End If
If glbCompSerial = "S/N - 2259W" Then '#15908
    lblStatus.Caption = "Income Code"
    clpCode(1).TABLTitle = "Income Code"
    vbxTrueGrid.Columns(3).Caption = "Income Code"
End If

'City of Niagara Falls - Ticket #16071
If glbCompSerial = "S/N - 2276W" Then
    lblHrsDay.Caption = "Hours/Pay Period"
End If

If glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    lblStatus.FontBold = False
    lblGroup.FontBold = False
End If
If glbCompSerial = "S/N - 2418W" Or glbCompSerial = "S/N - 2433W" Then
    'Ticket #17280 Charton-Hobbs Inc
    'Ticket #21683 Kerry's Place Autism Services
    lblGroup.FontBold = False
End If

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    lblMachine.Visible = True
    clpCode(0).Visible = True
    cmdAttachJobFiles.Visible = True
    
    If Not gSec_Inq_Job_Files_Attachment Then
        cmdAttachJobFiles.Enabled = False
    End If
End If

If Not EERetrieve() Then Exit Sub

If glbCompDecHR = 3 Then
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        medPayScale(X).Format = "#,##0.000;(#,##0.000)"
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'medPayScale(X + 11).Format = "#,##0.000;(#,##0.000)"
        'medPayScale(X + 15).Format = "#,##0.000;(#,##0.000)"
        medPayScale(X + 20).Format = "#,##0.000;(#,##0.000)"
    Next X
End If
If glbCompDecHR = 4 Then
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        medPayScale(X).Format = "#,##0.0000;(#,##0.0000)"
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'medPayScale(X + 11).Format = "#,##0.0000;(#,##0.0000)"
        'medPayScale(X + 15).Format = "#,##0.0000;(#,##0.0000)"
        medPayScale(X + 20).Format = "#,##0.0000;(#,##0.0000)"
    Next X
End If

If Not EERetrieve() Then Exit Sub
'If glbWFC Then clpCode(6).Left = 1300

lblComRatio.Visible = Not glbWFC
cmbMidPoint.Visible = Not glbWFC
comPayPer.Visible = Not glbWFC
lblSalary.Visible = Not glbWFC
fraGrid.Visible = Not glbWFC

Call combMidPoint

comPayPer.AddItem "Annual Grid"
comPayPer.AddItem "Hourly Grid"
'Data1.Refresh

If glbLinamar Then
    For X = 4 To 5
        lblWF(X - 3).Visible = True
         clpCode(X).Visible = True
         clpCode(X).ShowDescription = True
    Next X
    lblLGroup.Visible = True
    clpLGroup.Visible = True
End If

Me.Show

If Val(medFTEHrs.Text) = 0 Then
    medFTEHrs = ""
End If

If glbCompSerial = "S/N - 2191W" Then clpCode(2).MaxLength = 6

Call modSTUPD(False)

If Not gSec_Upd_Job_Master Then        'May99 js
    Call set_Buttons
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call SetMultiGrid
Call INI_Controls(Me)
Call Display_Value

If glbWFC Then 'Ticket #25785 Franks 07/30/2014
    'Call PointsTextboxScreenSetup
    Call WFCV81ScreenSetup 'Ticket #25911 Franks 09/30/2014
End If

Screen.MousePointer = DEFAULT
                             '
End Sub

Private Sub WFCV81ScreenSetup() 'Ticket #25911 Franks 09/30/2014
    '"   Hide all fields from # of positions down and all grid steps for hourly employees. Don't hide POINTS.
    lblNoPos.Visible = False
    txtNoPos.Visible = False
    lblPosFilled.Visible = False
    lblPosFiled.Visible = False
    'lblFTENum.Visible = False
    'medFTENum.Visible = False
    'lblTotNum.Visible = False
    'lblFTETotNum.Visible = False
    'lblFTEHrs.Visible = False
    'medFTEHrs.Visible = False
    'lblTotHrs.Visible = False
    'lblFTETotHrs.Visible = False
    'lblHrsDay.Visible = False
    'medHours.Visible = False
    
    'Ticket #28340 Franks 03/21/2016 - hide FTE #, FTE Hours/Year, Hours/Day
    lblFTENum.Visible = False
    medFTENum.Visible = False
    lblTotNum.Visible = False
    lblFTETotNum.Visible = False
    lblFTEHrs.Visible = False
    medFTEHrs.Visible = False
    lblTotHrs.Visible = False
    lblFTETotHrs.Visible = False
    lblHrsDay.Visible = False
    medHours.Visible = False
    'Ticket #28340 Franks 03/21/2016 - end
    
    frmWFCDIV.BorderStyle = 0

    lblPlant.Caption = lStr("Section")
    lblDiv.Caption = lStr("Division")
    lblRegion.Caption = lStr("Region")
    lblUnion.Caption = lStr("Union")
    
    vbxTrueGrid.Columns(3).Caption = lStr("Division")
    vbxTrueGrid.Columns(4).Caption = lStr("Region")
    vbxTrueGrid.Columns(5).Caption = lStr("Section")
    vbxTrueGrid.Columns(11).Caption = lStr("Position Level")
    vbxTrueGrid.Columns(13).Caption = lStr("Union")
    
    lblStatus.Caption = lStr("Position Status")
    lblGroup.Caption = lStr("Position Group")
    lblLevel.Caption = lStr("Position Level")
    lblPosDesc.Caption = lStr("Position Description")
    lblPosAlter.Caption = lStr("Position Alternate")
    lblUserDef1.Caption = lStr("Position User Defined 1")
    lblUserDef2.Caption = lStr("Position User Defined 2")
    
    clpCode(10).TABLTitle = UCase(lStr("Position Level") & " Codes")
    
    clpReportsTo(0).TextBoxWidth = 1305
    clpReportsTo(1).TextBoxWidth = 1305
    clpReportsTo(2).TextBoxWidth = 1305
    clpReportsTo(3).TextBoxWidth = 1305
End Sub

Private Sub PointsTextboxScreenSetup()
    lblPoints.DataField = ""
    lblPoints.Visible = False
    medPoints.Left = medHours.Left
    medPoints.DataField = "JB_POINTS"
    medPoints.Visible = True
End Sub

Private Sub SetMultiGrid()
Dim X
    If glbMultiGrid Then
        lblComRatio.Visible = False
        cmbMidPoint.Visible = False
        comPayPer.Visible = False
        lblSalary.Visible = False
        fraGrid.Visible = False
        clpCode(3).Visible = False
        lblUnion.Visible = False
        lblNoPos.Visible = False
        txtNoPos.Visible = False
        lblPosFiled.Visible = False
        lblPosFilled.Visible = False
        lblFTENum.Visible = False
        medFTENum.Visible = False
        lblTotNum.Visible = False
        lblFTETotNum.Visible = False
        lblFTEHrs.Visible = False
        medFTEHrs.Visible = False
        lblTotHrs.Visible = False
        lblFTETotHrs.Visible = False
        'lblTotalPoints.Visible = False
        'lblPoints.Visible = False
        lblCountWarn.Visible = False
        'cmdCountPos.Visible = False
        
        lblReptAuthor(0).Top = 1350
        clpReportsTo(0).Top = 1350
        lblReptAuthor(1).Top = 1680
        clpReportsTo(1).Top = 1680
        lblReptAuthor(2).Top = 2000
        clpReportsTo(2).Top = 2000
        lblProv.Top = 2310
        clpNationalClass.Top = 2310
        lblTotalPoints.Top = 2620
        lblPoints.Top = 2620
        
        lblHrsDay.Top = 2930
        medHours.Top = 2930
        
        lblGridC.Top = 2930 + 310
        clpGrid.Top = 2930 + 310
        txtLambtonJob.Top = 2930 + 310
        lblLambtonJob.Top = 2930 + 310
        lblLambtonJob.Left = 6330
        txtLambtonJob.Left = 7860
        
        cmdCountPos.Caption = "&Count Total Points"
        vbxTrueGrid.Columns(5).Visible = False
        
        For X = 11 To vbxTrueGrid.Columns.count - 1
            vbxTrueGrid.Columns(X).Visible = False
        Next
    End If
    
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
fraDetail.Height = 7860 '6300 '5000
If glbLinamar Then fraDetail.Height = 5700

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= vbxTrueGrid.Height + fraDetail.Height + panControls.Height + 730 Then
        scrControl.Value = 0
        fraDetail.Top = vbxTrueGrid.Height + 240
        scrControl.Visible = False
    Else
        'If Me.Height < vbxTrueGrid.Height + scrControl.Top + panControls.Height + 730 Then Exit Sub
        scrControl.Visible = True
        scrControl.Max = vbxTrueGrid.Height + fraDetail.Height + panControls.Height + 700 - Me.Height
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - panControls.Height - 550 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 550
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'fraDetail.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9300 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 170
        Else
            scrHScroll.Max = 30
        End If
        'scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Unload frmPosEval  'added by RAUBREY 5/23/97
'Unload frmPosSkills 'added by RAUBREY 5/23/97
Set frmMPOSITIONSWFC = Nothing   'carmen may 2000

End Sub

Private Sub imgIDiv_Click()
Call txtJobCode_DblClick
End Sub

Private Sub lblSalCode_Change()
If flagLoad = 0 Then Exit Sub    'carmen may 2000

If Not IsNull(lblSalCode.Caption) Then
    If lblSalCode.Caption = "A" Then
        comPayPer.ListIndex = 0
    ElseIf lblSalCode.Caption = "H" Then
        comPayPer.ListIndex = 1
    End If
End If

End Sub


Private Sub medFTEHrs_Change()

If medFTENum.Text = "" Or medFTENum.Text = "0" Then
    txtNoPos.Enabled = True 'And cmdOK.Enabled
Else
    txtNoPos.Enabled = False
End If

End Sub

Private Sub medFTEHrs_GotFocus()

medFTEHrs.MaxLength = 7
Call SetPanHelp(ActiveControl)

End Sub

Private Sub medFTEHrs_KeyPress(KeyAscii As Integer)

If medFTEHrs.Text = "" Then
    txtNoPos.Enabled = True 'And cmdOK.Enabled
Else
    txtNoPos.Enabled = False
End If

End Sub

Private Sub medFTEHrs_LostFocus()
    'Hemu - Ticket #10139 - Town of Aurora only
    'Oxford Ticket #15590
    If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2259W" Then
        Call Calculate_Secondary_Grid_Steps(0)
    End If
    
    'Ticket #25469 - City of Campbell River
    If glbCompSerial = "S/N - 2458W" And IsNumeric(medFTEHrs) And Not IsNumeric(medHours) Then
        'Compute Hours/Day
        medHours = medFTEHrs / 260
    End If
End Sub

Private Sub Calculate_Secondary_Grid_Steps(xStepNo As Integer)

If xStepNo = 0 Then
    If medFTEHrs.Text <> "" And medFTEHrs.Text <> "0" Then
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'For xStep = 1 To 11
        'For xStep = 1 To 15
        For xStep = 1 To 20
            If lblSalCode.Caption = "A" Then
                'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
                'medPayScale(xStep + 11).Text = Val(medPayScale(xStep).Text) / Val(medFTEHrs.Text)
                'medPayScale(xStep + 15).Text = Val(medPayScale(xStep).Text) / Val(medFTEHrs.Text)
                medPayScale(xStep + 20).Text = Val(medPayScale(xStep).Text) / Val(medFTEHrs.Text)
            ElseIf lblSalCode.Caption = "H" Then
                'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
                'medPayScale(xStep + 11).Text = Val(medPayScale(xStep).Text) * Val(medFTEHrs.Text)
                'medPayScale(xStep + 15).Text = Val(medPayScale(xStep).Text) * Val(medFTEHrs.Text)
                medPayScale(xStep + 20).Text = Val(medPayScale(xStep).Text) * Val(medFTEHrs.Text)
            End If
        Next xStep
    Else
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'For xStep = 1 To 11
        'For xStep = 1 To 15
        For xStep = 1 To 20
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'medPayScale(xStep + 11).Text = ""
            'medPayScale(xStep + 15).Text = ""
            medPayScale(xStep + 20).Text = ""
        Next xStep
    End If
Else
    If medFTEHrs.Text <> "" And medFTEHrs.Text <> "0" Then
        If lblSalCode.Caption = "A" Then
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'medPayScale(xStepNo + 11).Text = Val(medPayScale(xStepNo).Text) / Val(medFTEHrs.Text)
            'medPayScale(xStepNo + 15).Text = Val(medPayScale(xStepNo).Text) / Val(medFTEHrs.Text)
            medPayScale(xStepNo + 20).Text = Val(medPayScale(xStepNo).Text) / Val(medFTEHrs.Text)
        ElseIf lblSalCode.Caption = "H" Then
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'medPayScale(xStepNo + 11).Text = Val(medPayScale(xStepNo).Text) * Val(medFTEHrs.Text)
            'medPayScale(xStepNo + 15).Text = Val(medPayScale(xStepNo).Text) * Val(medFTEHrs.Text)
            medPayScale(xStepNo + 20).Text = Val(medPayScale(xStepNo).Text) * Val(medFTEHrs.Text)
        End If
    Else
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'medPayScale(xStepNo + 11).Text = ""
        'medPayScale(xStepNo + 15).Text = ""
        medPayScale(xStepNo + 20).Text = ""
    End If
End If

End Sub

Private Sub medFTENum_Change()

If medFTENum.Text = "" Or medFTENum.Text = "0" Then
    txtNoPos.Enabled = True 'And cmdOK.Enabled
Else
    txtNoPos.Enabled = False
End If

End Sub

Private Sub medFTENum_GotFocus()

medFTENum.MaxLength = 6 'allows for 6 keystrokes including "."
Call SetPanHelp(ActiveControl)

End Sub

Private Sub medFTENum_KeyPress(KeyAscii As Integer)

If medFTENum.Text = "" Then
    txtNoPos.Enabled = True 'And cmdOK.Enabled
Else
    txtNoPos.Enabled = False
End If

End Sub

Private Sub medHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPayScale_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub modChekCount()

On Error GoTo modChekCount_Err

lblCountWarn.Visible = False

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If Not IsNull(Data1.Recordset("JB_NBRPOS")) Then
        If Not IsNull(Data1.Recordset("JB_NBRFIL")) Then
            If CInt(Data1.Recordset("JB_NBRPOS")) < CInt(Data1.Recordset("JB_NBRFIL")) Then
                 lblCountWarn.Visible = True
            End If
        End If
    End If
End If

Exit Sub

modChekCount_Err:

End Sub

Private Function modJobSectionUnique(Job, JobID, xSection)
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset
modJobSectionUnique = True
SQLQ = "SELECT JB_CODE FROM HRJOB WHERE JB_CODE='" & Trim(Job) & "' AND JB_ID<>" & JobID & " "
SQLQ = SQLQ & "AND JB_SECTION = '" & xSection & "' "
rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsJOB.EOF Then modJobSectionUnique = False
rsJOB.Close

End Function

Private Function modJobUnique(Job, JobID)
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset
modJobUnique = True
rsJOB.Open "SELECT JB_CODE FROM HRJOB WHERE JB_CODE='" & Trim(Job) & "' AND JB_ID<>" & JobID, gdbAdoIhr001, adOpenForwardOnly
If Not rsJOB.EOF Then modJobUnique = False
rsJOB.Close
End Function

Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer, X

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If
clpGrid.Visible = False
lblGridC.Visible = False
txtLambtonJob.Visible = False
lblLambtonJob.Visible = False
  
'cmdOK.Enabled = TF          'May99 js
'cmdCancel.Enabled = TF      '
cmdCountPos.Enabled = TF  'FT    '
'cmdNew.Enabled = FT         '
'cmdFind.Enabled = TF
'cmdClose.Enabled = FT       '
'cmdModify.Enabled = FT      '
'cmdDelete.Enabled = FT      '
'cmdPrint.Enabled = FT       '
'vbxTrueGrid.Enabled = FT
'cmbBand.Enabled = TF        'Jaddy 8/9/99
clpCode(6).Enabled = TF
cmbMidPoint.Enabled = TF    '
'cmdCountPos.Enabled = TF    '
comPayPer.Enabled = TF      '
fraGrid.Enabled = TF        '
medFTEHrs.Enabled = TF      '
medFTENum.Enabled = TF      '
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For X = 1 To 11
'For X = 1 To 15
For X = 1 To 20
    medPayScale(X).Enabled = TF '
Next
 clpCode(1).Enabled = TF     '
 clpCode(2).Enabled = TF     '
 clpCode(3).Enabled = TF     '
 clpCode(4).Enabled = TF     '
 clpCode(5).Enabled = TF     '
 clpNationalClass.Enabled = TF '
txtNoPos.Enabled = TF       '
txtPosDescr.Enabled = TF    '
txtPosDescr2.Enabled = TF   '
 clpReportsTo(0).Enabled = TF   '
 clpReportsTo(1).Enabled = TF   '
 clpReportsTo(2).Enabled = TF   '
txtPosition.Enabled = FT   ' dkostka - 03/06/01 - Jerry requested

'George on Feb 3,2006 #10266
glbDocName = "Jobdescription"
If gsAttachment_DB Then
    'glbPos = Data1.Recordset("JB_CODE")
    Call DispimgIcon(Me, "frmMPOSITIONS")
    If gSec_Upd_Job_Master Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False 'George on Jan 26,2006 #10266
        Else
            cmdImport.Visible = True 'George on Jan 26,2006 #10266
        End If
    End If
End If
'George on Feb 3,2006 #10266

clpLGroup.Enabled = TF
End Sub

Private Sub medPayScale_LostFocus(Index As Integer)
    'Hemu - Ticket #10139 - Town of Aurora only
    'Oxford Ticket #15590
    If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2259W" Then
        Call Calculate_Secondary_Grid_Steps(Index)
    End If
End Sub

Private Sub medPoints_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
fraDetail.Top = 70 + vbxTrueGrid.Height - scrControl.Value * 1.2
End Sub

Private Sub txtBand_Change()
  'If glbWFC And (clpCode(3).Text = "NONE" Or clpCode(3).Text = "EXEC") Then cmbBand_SETUP Me
End Sub
Private Sub clpCode_Change(Index As Integer)
    If Index = 3 Then
        If glbWFC Then
            If clpCode(3) = "NONE" Or clpCode(3) = "EXEC" Then
                lblComRatio.Visible = False
                cmbMidPoint.Visible = False
                comPayPer.Visible = False
                lblSalary.Visible = False
                fraGrid.Visible = False
                'cmbBand.Visible = True
                lblBand.Visible = True
                clpCode(6).Visible = True
            Else
                'cmbBand.Visible = False
                lblBand.Visible = False
                clpCode(6).Visible = False
                ''Ticket #25911 Franks 10/01/2014 - begin
                'lblComRatio.Visible = True
                'cmbMidPoint.Visible = True
                'comPayPer.Visible = True
                'lblSalary.Visible = True
                'fraGrid.Visible = True
                lblComRatio.Visible = False
                cmbMidPoint.Visible = False
                comPayPer.Visible = False
                lblSalary.Visible = False
                fraGrid.Visible = False
                ''Ticket #25911 Franks 10/01/2014 - end
            End If
        End If
        comPayPer.Refresh
    End If
    If Index = 12 Then 'Ticket #29552 Franks 12/14/2016
            Call EERetrieve
    End If
End Sub

Private Sub scrHScroll_Change()
fraDetail.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtJobCode_Change()
    lblJobCodeDesc.Caption = getJobCodeDesc(txtJobCode.Text)
    
    'Ticket #30313 Franks 07/03/2017 - begin
    If txtJobCode.Text = "IND000" Then
        lblReptAuthor(0).FontBold = False
        lblProv.FontBold = False
        lblPosType.FontBold = False
    Else
        lblReptAuthor(0).FontBold = True
        lblProv.FontBold = True
        lblPosType.FontBold = True
    End If
    'Ticket #30313 Franks 07/03/2017 - end
    
End Sub
Private Function getJobCodeDesc(xCode)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = "Unassigned"
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("JB_JOBDESCR")
        End If
        rsDiv.Close
    End If
    getJobCodeDesc = xRetVal
    If Len(xCode) > 0 Then
        lblJobCodeDesc.Visible = True
    Else
        lblJobCodeDesc.Visible = False
    End If
End Function
Private Sub txtJobCode_DblClick()
    Call Get_JobMaster(False)
    If Len(glbJobMaster) > 0 Then
        txtJobCode.Text = glbJobMaster
    End If
End Sub

Private Sub txtJobCode_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtJobCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtJobCode_LostFocus()
    If glbWFC And fglbNew Then
        If Len(txtJobCode.Text) > 0 Then
            Call getDataFromJobMaster(txtJobCode.Text)
        End If
    End If
End Sub

Private Sub getDataFromDivMaster(xCode)
Dim SQLQ As String
Dim rs As New ADODB.Recordset
    If Len(xCode) > 0 Then
        SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xCode & "' "
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("DV_REGION")) Then clpCode(9).Text = rs("DV_REGION")
            If Not IsNull(rs("DV_SECTION")) Then clpCode(7).Text = rs("DV_SECTION")
            If Not IsNull(rs("DV_MARKETLINE")) Then clpCode(8).Text = rs("DV_MARKETLINE") 'Market Line
        End If
        rs.Close
        If Not (xCode = "1000" Or xCode = "1001" Or xCode = "8000" Or xCode = "8001") Then 'Ticket #29220 Franks 09/19/2016
            clpCode(11).Text = "PLANT"
        End If
    End If
End Sub

Private Sub getDataFromJobMaster(xCode)
Dim SQLQ As String
Dim rs As New ADODB.Recordset
    If Len(xCode) > 0 Then
        SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & xCode & "' "
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("JB_STATUS")) Then clpCode(1).Text = rs("JB_STATUS")
            If rs("JB_STATUS") = "CLER" Or rs("JB_STATUS") = "MGMT" Or rs("JB_STATUS") = "SUPR" Then 'Ticket #29183 Franks 08/14/2016
                clpCode(3).Text = "NONE"
            End If
            If Not IsNull(rs("JB_GRPCD")) Then clpCode(2).Text = rs("JB_GRPCD")
        End If
        rs.Close
    End If
End Sub

Private Sub txtMercerNo_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMidPoint_Change()
    cmbMidPoint.ListIndex = 0
    If IsNumeric(txtMidPoint) Then cmbMidPoint.ListIndex = Val(txtMidPoint) - 1
End Sub

Private Sub txtNoPos_Change()

If txtNoPos.Text = "" Or txtNoPos.Text = "0" Then
    medFTENum.Enabled = True 'And cmdOK.Enabled
    medFTEHrs.Enabled = True 'And cmdOK.Enabled
Else
    medFTENum.Enabled = False
    
    'Vadim uses this field to compute the Hourly Salary to transfer to iCity from the Annual Grid.
    If Not glbVadim Then
        medFTEHrs.Enabled = False
    End If
End If

If medFTENum.Text = "" Or medFTENum.Text = "0" Then
    txtNoPos.Enabled = True 'And cmdOK.Enabled
Else
    txtNoPos.Enabled = False
End If

End Sub

Private Sub txtNoPos_Click()

If txtNoPos.Text = "" Or txtNoPos.Text = "0" Then
    medFTENum.Enabled = True 'And cmdOK.Enabled
Else
    medFTENum.Enabled = False
End If

End Sub

Private Sub txtNoPos_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtNoPos_KeyPress(KeyAscii As Integer)

If txtNoPos.Text = "" Or txtNoPos.Text = "0" Then
    medFTENum.Enabled = True 'And cmdOK.Enabled
Else
    medFTENum.Enabled = False
End If

End Sub

Private Sub txtPosDescr_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPosDescr2_Change()
Call combMidPoint
End Sub

Private Sub txtPosDescr2_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPosition_Change()

If txtPosition.Enabled = False Then
    Me.Caption = "Position Information - " & txtPosition
End If

End Sub

Private Sub txtPosition_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtPosition_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
'Private Sub txtReportsTo_Change(Index As Integer)
'If flagLoad = 0 Then Exit Sub
'    Call Job_Desc(Index)
'End Sub
'Private Sub txtReportsTo_DblClick(Index As Integer)
'Dim OJOB As String, OJobD As String
'OJOB = txtReportsTo(Index).Text
'OJobD = lblJobDesc(Index).Caption
'Load frmJOBS
'frmJOBS.Show 1
'If Len(glbJob) < 1 Then
'    txtReportsTo(Index).Text = OJOB
'    lblJobDesc(Index).Caption = OJobD
'Else
'    txtReportsTo(Index).Text = glbJob
'    lblJobDesc(Index).Caption = glbJobDesc
'End If
'End Sub
'Private Sub txtReportsTo_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtReportsTo_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 8 Then    ' hit backspace
'    Exit Sub
'End If
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Private Sub txtPosition_LostFocus()
    If txtPosition <> "" And clpGrid <> "" Then
        txtLambtonJob = Left(clpGrid, 1) & txtPosition & Mid(clpGrid, 2, 1)
    End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
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
        
        SQLQ = "SELECT * FROM HRJOB WHERE 1 = 1"
        If Len(glbWFCUserSecList) > 0 Then 'Ticket #27609 Franks 10/13/2015
            SQLQ = SQLQ & " AND JB_SECTION IN " & glbWFCUserSecList & " "
        End If
        If Len(clpCode(12).Text) > 0 Then 'Ticket #29552 Franks 12/14/2016
            SQLQ = SQLQ & "AND JB_SECTION = '" & clpCode(12).Text & "' "
        End If
        If chkHideInactive Then
            SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
            If glbOracle Then 'Ticket #16416
                SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
            ElseIf glbSQL Then
                SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
            Else
                SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
            End If
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Me.vbxTrueGrid.EditActive Then
'        cmdOK.SetFocus
    Else
 '       cmdClose.SetFocus
    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X As Integer

On Error GoTo RCErr
'Call Display_Value

Call modChekCount
' dkostka = 08/31/2000 - Made change on request by glbWFC, hide salary grids only if union is NONE or EXEC
'Jaddy move from date1_movecompled
If glbWFC Then
    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
        If Data1.Recordset("JB_ORG") = "NONE" Or Data1.Recordset("JB_ORG") = "EXEC" Then
            lblComRatio.Visible = False
            cmbMidPoint.Visible = False
            comPayPer.Visible = False
            lblSalary.Visible = False
            fraGrid.Visible = False
            'cmbBand.Visible = True
            lblBand.Visible = True
            clpCode(6).Visible = True
        Else
            'cmbBand.Visible = False
            lblBand.Visible = False
            clpCode(6).Visible = False
            ''Ticket #25911 Franks 10/01/2014 - begin
            ''"   For hourly employees, always hide the Salary Grid and Step #.
            'lblComRatio.Visible = True
            'cmbMidPoint.Visible = True
            'comPayPer.Visible = True
            'lblSalary.Visible = True
            'fraGrid.Visible = True
            lblComRatio.Visible = False
            cmbMidPoint.Visible = False
            comPayPer.Visible = False
            lblSalary.Visible = False
            fraGrid.Visible = False
            ''Ticket #25911 Franks 10/01/2014 - end
        End If
        comPayPer.Refresh
    End If
End If

If Data1.Recordset.EOF Then
    glbPos = ""
    glbPosDesc = ""
    glbJobSection = ""
Else
    glbPos = Data1.Recordset("JB_CODE")
    glbPosDesc = Data1.Recordset("JB_DESCR")
    If IsNull(Data1.Recordset("JB_SECTION")) Then glbJobSection = "" Else glbJobSection = Data1.Recordset("JB_SECTION")
End If


If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If Data1.Recordset("JB_SALCD") = "A" Then
        comPayPer.ListIndex = 0
    ElseIf Data1.Recordset("JB_SALCD") = "H" Then
        comPayPer.ListIndex = 1
    Else
        comPayPer.ListIndex = -1
    End If
End If

Call Display_Value

oStatus = clpCode(1).Text

'close these forms
Unload frmPosEval  'added by RAUBREY 5/23/97
Unload frmPosSkills 'added by RAUBREY 5/23/97

Exit Sub

RCErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Row change", "HRJob", "SELECT")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
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

Private Sub GridStepChange()
Dim fTablSalHis As New ADODB.Recordset
Dim lngLastCurrentID&
Dim Msg$, SQLQ, Response%
Dim X%, I
Dim xStr
Dim IfChangeMatch As Boolean
Dim Emp_List As New Collection
Dim num
Dim xSHID
Dim xPreStep 'Ticket #20076
Dim oldRate, newRate

    On Error GoTo ErrorHandler
    
    strEMPLIST = ""
    strEmpEffError = ""
    
    'If fglbNew Then
    '    GoTo MarkExit0
    'End If

    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        DoEvents
        GridStep(X%, 1) = medPayScale(X%).Text
        
        'Hemu - Town of Aurora only - Ticket #10263
        If glbCompSerial = "S/N - 2378W" Then
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'GridStep2(X%, 1) = medPayScale(X% + 11).Text
            'GridStep2(X%, 1) = medPayScale(X% + 15).Text
            GridStep2(X%, 1) = medPayScale(X% + 20).Text
        End If
        
        DoEvents
        
        If glbFrench Then
            If GridStep(X%, 0) <> CDbl(GridStep(X%, 1)) Then
                GridStep(X%, 2) = "U"
                
                'Hemu - Town of Aurora only - Ticket #10263
                If glbCompSerial = "S/N - 2378W" Then
                    GridStep2(X%, 2) = "U"
                End If
                
                IfGridStepChange = True
                If glbVadim And Not glbMultiGrid Then
                    'Hemu - Town of Aurora only - Ticket #10263
                    If glbCompSerial = "S/N - 2378W" Then
                        If Left(comPayPer.Text, 1) = "A" Then
                            Call Passing_Salary_Grid_Vadim(X%, GridStep2(X%, 0), GridStep2(X%, 1), Date, txtPosition)
                        Else
                            Call Passing_Salary_Grid_Vadim(X%, GridStep(X%, 0), GridStep(X%, 1), Date, txtPosition)
                        End If
                    Else
                        Call Passing_Salary_Grid_Vadim(X%, GridStep(X%, 0), GridStep(X%, 1), Date, txtPosition)
                    End If
                End If
            End If
        Else
            If GridStep(X%, 0) <> Val(GridStep(X%, 1)) Then
                GridStep(X%, 2) = "U"
                
                'Hemu - Town of Aurora only - Ticket #10263
                If glbCompSerial = "S/N - 2378W" Then
                    GridStep2(X%, 2) = "U"
                End If
                
                IfGridStepChange = True
                If glbVadim And Not glbMultiGrid Then
                    'Hemu - Town of Aurora only - Ticket #10263
                    If glbCompSerial = "S/N - 2378W" Then
                        If Left(comPayPer.Text, 1) = "A" Then
                            Call Passing_Salary_Grid_Vadim(X%, GridStep2(X%, 0), GridStep2(X%, 1), Date, txtPosition)
                        Else
                            Call Passing_Salary_Grid_Vadim(X%, GridStep(X%, 0), GridStep(X%, 1), Date, txtPosition)
                        End If
                    Else
                        'Ticket #25469 - City of Campbell River - Transfer Annual as Annual and Hourly as Hourly
                        'Ticket #23795 - Town of Lasalle - Annual to transfer as Annual and Hourly to transfer as Rate
                        'Ticket #21124 - Convert Annual amount to Hourly before passing to Vadim
                        If Left(comPayPer.Text, 1) = "A" And glbCompSerial <> "S/N - 2379W" And glbCompSerial <> "S/N - 2458W" Then
                        '    If xWHRS <> 0 Then oldRate = Round2DEC((GridStep(X%, 0) / 52) / xWHRS)
                        '    If xWHRS <> 0 Then newRate = Round2DEC((GridStep(X%, 1) / 52) / xWHRS)
                        '
                            'Ticket #23049 - City of Niagara Falls
                            If glbCompSerial = "S/N - 2276W" Then
                                'Do not transfer Hourly Rate from here to Vadim function cause they have custom
                                'formula to compute in hours in the Vadim function below.
                                oldRate = GridStep(X%, 0)
                                newRate = GridStep(X%, 1)
                            Else
                                If IsNumeric(medFTEHrs.Text) And Val(medFTEHrs.Text) <> 0 Then
                                    oldRate = Round2DEC(Val((GridStep(X%, 0)) / Val(medFTEHrs.Text)))
                                    newRate = Round2DEC(Val((GridStep(X%, 1)) / Val(medFTEHrs.Text)))
                                Else
                                    oldRate = 0
                                    newRate = 0
                                End If
                            End If
                            Call Passing_Salary_Grid_Vadim(X%, oldRate, newRate, Date, txtPosition)
                        Else
                            Call Passing_Salary_Grid_Vadim(X%, GridStep(X%, 0), GridStep(X%, 1), Date, txtPosition)
                        End If
                    End If
                End If
            End If
        End If
    Next X
    
    If Not IfGridStepChange Then
        'Hemu - To update the Comp Ratio if there is a change in the Mid Point
        '       Ticket #6471
        If lstMidPoint <> txtMidPoint Then
            Call Update_CompRatio
        End If
    
        GoTo MarkExit0
    Else
        'Hemu - To update the Comp Ratio (Ticket #6471)
        Call Update_CompRatio
        
        'City of Kawartha Lakes - They are getting deadlock error so this is to see if it works in their
        'test environment - Ticket #20321
        'If glbCompSerial = "S/N - 2363W" Then
        '    fTablSalHis.Open "HR_SALARY_HISTORY", gdbAdoIhr001, adOpenKeyset, adLockPessimistic
        'Else
            fTablSalHis.Open "HR_SALARY_HISTORY", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'End If
        If glbOracle Then
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY,HRJOB WHERE HRJOB.JB_CODE = HR_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & " AND SH_CURRENT <> 0 and SH_JOB = '" & txtPosition & "'"
        Else
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & " WHERE SH_CURRENT <> 0 and SH_JOB = '" & txtPosition & "'"
        End If
        dynSH_Job1.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        If dynSH_Job1.EOF And dynSH_Job1.BOF Then
            GoTo MarkExit
        End If

        Msg$ = "Do you want to update the employee salaries too?"
        X = MsgBox(Msg, 36, "Confirm Update")
        If X <> 6 Then GoTo MarkExit
        
        If CheckGridStepChange = False Then GoTo MarkExit
        
        strEMPLIST = ""
        strEmpEffError = ""
    End If
    
    glbGridReason = ""
    glbGridEDate = ""
    glbGridNDate = ""
    Load frmForGridStep
    frmForGridStep.Show 1

    If Len(glbGridReason) = 0 Then GoTo MarkExit
    
    Msg$ = "Do you want to print a list of employees updated?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Position Master - Salary Update")      ' Get user response.
        
    
    Screen.MousePointer = HOURGLASS
    
    '---- main function
    dynSH_Job1.MoveFirst
    Do While Not dynSH_Job1.EOF
         Emp_List.Add (dynSH_Job1("SH_ID"))
         dynSH_Job1.MoveNext
    Loop
    
    'Do While Not dynSH_Job1.EOF
    DoEvents
    For num = 1 To Emp_List.count
        dynSH_Job1.MoveFirst
        
        lngLastCurrentID& = Emp_List(num) 'dynSH_Job1("SH_ID")
        dynSH_Job1.Find "SH_ID = " & lngLastCurrentID&
        
        DoEvents
        
        empNo& = dynSH_Job1("SH_EMPNBR")
        xStr = dynSH_Job1("SH_GRADE")
        
        ' dkostka - 01/28/2002 - Changed below to interpret null grade/step as 00.  A code problem
        '   in a previous release caused this, but has since been fixed (but DB may still have nulls).
        If IsNull(xStr) Then xStr = "00"
                
        IfChangeMatch = False
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'For X% = 1 To 11
        'For X% = 1 To 15
        For X% = 1 To 20
            DoEvents
            If GridStep(X%, 2) = "U" Then ' if update
                If X% = Val(xStr) Then
                    ' this employee's SH_GRADE match one of the changed Grid Step
                    IfChangeMatch = True
                    GoTo Mark1
                End If
            End If
        Next X%
        If Not IfChangeMatch Then GoTo NextRec
Mark1:
        'lngLastCurrentID& = dynSH_Job1("SH_ID")
        fTablSalHis.MoveFirst
        fTablSalHis.Find "SH_ID = " & lngLastCurrentID&
        
        OSalary = fTablSalHis("SH_SALARY")
        OTOTAL = fTablSalHis("SH_TOTAL")
        oPayP = fTablSalHis("SH_PAYP")
        OEDate = fTablSalHis("SH_EDATE")
        oPayrollID = fTablSalHis("SH_PAYROLL_ID")
        ONDate = fTablSalHis("SH_NEXTDAT")
        OJOB1 = fTablSalHis("SH_JOB")
        OSalCD = fTablSalHis("SH_SALCD")
        oGrade = fTablSalHis("SH_GRADE")
        
        If Len(fTablSalHis("SH_WHRS")) < 1 Then
            dblWHours# = 0
        Else
            dblWHours# = fTablSalHis("SH_WHRS")
        End If
        
        'Ticket #20588
        'Do not add a new salary if the new Effective Date is less than the current Salary Effective Date
            'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
            'update on the salary screen. So changed from >= to >.
        If OEDate > CVDate(glbGridEDate) Then
            'List of Employees not updated
            If Len(strEmpEffError) > 0 Then
                strEmpEffError = strEmpEffError & "," & dynSH_Job1("SH_EMPNBR")
            Else
                strEmpEffError = dynSH_Job1("SH_EMPNBR")
            End If
                        
            GoTo NextRec
        End If
        
        fTablSalHis("SH_LTIME") = Time$    '"T" & I
        fTablSalHis("SH_CURRENT") = False
        
        If glbCompSerial = "S/N - 2259W" Then 'Ticket #20076
            If IsNull(fTablSalHis("SH_GRADE")) Then
                xPreStep = ""
            Else
                xPreStep = fTablSalHis("SH_GRADE")
            End If
        End If
        fTablSalHis.Update
                
        'Ticket #16991 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
        'remaining same, only the actual salary is changing and this table only stores the Rate Level
        'Comment enhacement - Ticket #16115
        'City of Niagara Falls - Ticket #15542
        'If glbVadim And glbCompSerial = "S/N - 2276W" Then
        '    'Update previous salary record in Vadim's HR_EMP_HIST table with End Date
        '    Call Update_VadimDB_HR_EMP_HISTORY(oPayrollID, OEDate, "", "", "", "M", DateAdd("d", -1, CVDate(glbGridEDate)))
        'End If
        
        fTablSalHis.AddNew
        fTablSalHis("SH_COMPNO") = dynSH_Job1("SH_COMPNO")
        fTablSalHis("SH_EMPNBR") = dynSH_Job1("SH_EMPNBR")
        fTablSalHis("SH_EDATE") = CVDate(glbGridEDate)
        fTablSalHis("SH_CURRENT") = True
        fTablSalHis("SH_SDATE") = dynSH_Job1("SH_SDATE")
        fTablSalHis("SH_SALCD") = dynSH_Job1("SH_SALCD")
        fTablSalHis("SH_PAYROLL_ID") = dynSH_Job1("SH_PAYROLL_ID")
        fTablSalHis("SH_WHRS") = dynSH_Job1("SH_WHRS")
        fTablSalHis("SH_PAYP") = dynSH_Job1("SH_PAYP")
        fTablSalHis("SH_PAYP_TABLE") = dynSH_Job1("SH_PAYP_TABLE")
        fTablSalHis("SH_SREAS_TABLE") = dynSH_Job1("SH_SREAS_TABLE")
        dblOSalary = dynSH_Job1("SH_SALARY")

        If glbFrench Then
            dblNewSalary = CDbl(GridStep(X%, 1))
        Else
            dblNewSalary = Val(GridStep(X%, 1))
        End If
        
        'Get employee Hours/Day
        fglbDhrs = GetJHData(dynSH_Job1("SH_EMPNBR"), "JH_DHRS", 0)
        
        'CHanged by Bryan 09/08/05 Ticket #9086
        'Sent Workhours to GetNewStepSalary so that hourly salary could be calculated
        Call GetNewStepSalary(dblNewSalary, X%, dynSH_Job1("SH_WHRS"))

        ' dkostka - 01/29/2002 - Added list of employees that couldn't be changed for grid step changes.
        If dblNewSalary = -1 Then
            ' Couldn't change salary (no WHRS).  Abort and add to the list of errors.
            EmpChgErrors = EmpChgErrors & dynSH_Job1("SH_EMPNBR") & vbCrLf
            fTablSalHis.CancelUpdate
            'Added by Bryan 09/08/05 Ticket#9086
            'Resets Current Salary if a new record couldn't be created
            fTablSalHis.MoveFirst
            fTablSalHis.Find "SH_ID = " & lngLastCurrentID&
            fTablSalHis("SH_LTIME") = Time$    '"T" & I
            fTablSalHis("SH_CURRENT") = True
            fTablSalHis.Update
            'end bryan
            GoTo NextRec
        End If
        
        dblNewSalary = Round2DEC(dblNewSalary)
        fTablSalHis("SH_SALARY") = dblNewSalary
        
        If IsDate(glbGridNDate) Then
            fTablSalHis("SH_NEXTDAT") = glbGridNDate
        Else
            If IsDate(ONDate) Then
                If CVDate(ONDate) > CVDate(glbGridEDate) Then
                    fTablSalHis("SH_NEXTDAT") = ONDate
                End If
            End If
        End If
        
        fTablSalHis("SH_JOB") = dynSH_Job1("SH_JOB")
        fTablSalHis("SH_JOB_ID") = dynSH_Job1("SH_JOB_ID")

        Call modSetCOMPA_GRADE(dblNewSalary) ' sets fglbCOMPA#, and fglbGRADE
        
        fTablSalHis("SH_COMPA") = Round(fglbCOMPA#, 2)
        If glbCompSerial = "S/N - 2259W" Then 'Ticket #17139 user will update step manually
            'Ticket #20076 Franks 03/31/2011,Oxford needs it from Previous current salary
            If Len(xPreStep) = 0 Then
                xPreStep = fglbGRADE$
            End If
            fTablSalHis("SH_GRADE") = xPreStep
        Else
            fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")
        End If

        fTablSalHis("SH_SREAS1") = glbGridReason ' clpCode(4).Text
        If dblOSalary <> 0 Then fTablSalHis("SH_SALPC1") = (dblNewSalary - dblOSalary) / dblOSalary
        fTablSalHis("SH_SALCHG1") = dblNewSalary - dblOSalary

        fTablSalHis("SH_TRANSDATE") = Now   'Ticket #22455
        fTablSalHis("SH_LDATE") = Now
        fTablSalHis("SH_LTIME") = Time$ '"T" & I
        fTablSalHis("SH_LUSER") = glbUserID
        If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
            fTablSalHis("SH_PREMIUM") = dynSH_Job1("SH_PREMIUM")
            fTablSalHis("SH_TOTAL") = fTablSalHis("SH_SALARY") + dynSH_Job1("SH_PREMIUM")
            fTablSalHis("SH_VGROUP") = dynSH_Job1("SH_VGROUP")
            fTablSalHis("SH_VSTEP") = dynSH_Job1("SH_VSTEP")
        End If
        fTablSalHis.Update
        xSHID = fTablSalHis("SH_ID")
        
        'List of Employees updated
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & dynSH_Job1("SH_EMPNBR")
        Else
            strEMPLIST = dynSH_Job1("SH_EMPNBR")
        End If

        If glbVadim Then Call Transfer_Salary(fTablSalHis)
        
        Call updFollow("U") 'If Next Review Date enter 0
        
        Call updBenefitForSalDEPN(empNo&)
        
        'Ticket #16991 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
        'remaining same, only the actual salary is changing and this table only stores the Rate Level
        'City of Niagara Falls - Ticket #15542
        'If glbVadim And glbCompSerial = "S/N - 2276W" Then
        '    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
        '    Call Update_VadimDB_HR_EMP_HISTORY(fTablSalHis("SH_PAYROLL_ID"), CVDate(glbGridEDate), "", Val(fglbGRADE$), fTablSalHis("SH_JOB"), "A")
        'End If
        
        'If Not glbWFC Then 'Greensboro
            Call Employee_Master_Integration(empNo&)
        'End If
        If glbGP Then Call Salary_Integration(empNo&, , False, True, xSHID)
        NSalary = dblNewSalary
        NEDate = CVDate(glbGridEDate) '(txtEDate) 1
        NNDate = ""
        If Not AUDITSALY() Then MsgBox "ERROR - AUDIT FILE"
        
NextRec:
        'dynSH_Job1.MoveNext
    'Loop
    Next
    
    
'MsgBox "step 4 "
MarkExit:
    dynSH_Job1.Close
    fTablSalHis.Close

    
    If Response% = IDYES Then    ' Yes response
    
        Screen.MousePointer = DEFAULT
        
        'Employees Updated
        If Len(strEMPLIST) > 0 Then
            MsgBox "Salary Records Updated Successfully."
        Else
            MsgBox "0 Salary Records Updated."
        End If
            
        Screen.MousePointer = HOURGLASS
        
        'report name
        Me.vbxCrystal1.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    
        Me.vbxCrystal1.Formulas(0) = "rTitle='Position Master - Salary Update'"
        Me.vbxCrystal1.Formulas(1) = "EmpGroup='Employee(s) Updated:'"
        
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            If Len(strEMPLIST) > 0 Then
                Me.vbxCrystal1.SelectionFormula = getWSQLQRPT(strEMPLIST)
            Else
                Me.vbxCrystal1.SelectionFormula = "1=2"
            End If
        End If
        Me.vbxCrystal1.Connect = RptODBC_SQL
        
        'window title if appropriate
        Me.vbxCrystal1.WindowTitle = "Employee(s) Updated Report"
        
        Me.vbxCrystal1.Destination = 0
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal1.Action = 1
        vbxCrystal1.Reset
        
        'Employees Not Updated
        If Len(strEmpEffError) > 0 Then
            Me.vbxCrystal1.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
            Me.vbxCrystal1.Formulas(0) = "rTitle='Position Master - Salary Update'"
            Me.vbxCrystal1.Formulas(1) = "EmpGroup='Employee(s) Not Updated:'"
            
            'set location for database tables
            If Len(glbstrSelCri) >= 0 Then
                If Len(strEmpEffError) > 0 Then
                    Me.vbxCrystal1.SelectionFormula = getWSQLQRPT(strEmpEffError)
                Else
                    Me.vbxCrystal1.SelectionFormula = "1=2"
                End If
            End If
            Me.vbxCrystal1.Connect = RptODBC_SQL
            
            'window title if appropriate
            Me.vbxCrystal1.WindowTitle = "Employee(s) Not Updated Report"
            
            Me.vbxCrystal1.Destination = 0
            Screen.MousePointer = DEFAULT
            Me.vbxCrystal1.Action = 1
            vbxCrystal1.Reset
        End If
    End If

MarkExit0:
    Screen.MousePointer = DEFAULT

    Exit Sub
        
ErrorHandler:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Changing Grid Steps", "GRID_STEPS", "UPDATE")
    If False Then
        ' for debugging
        Resume
    End If
End Sub

Private Sub GetNewStepSalary(dblNewSalary, X%, dblHoursPerWeek#)
 
If Data1.Recordset("JB_SALCD") = "H" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##"))
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        If dblHoursPerWeek# = 0 Then
            dblNewSalary = -1
        Else
            dblNewSalary = (Data1.Recordset("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)) / 12
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        If dblHoursPerWeek# = 0 Then
            dblNewSalary = -1
        Else
            dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)
        End If
    'Day added by Bryan 28/Sep/05 Ticket#9354
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If dblHoursPerWeek# = 0 Then
            dblNewSalary = -1
        Else
            If GetLeapYear(Year(Date)) Then
                dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52) / 366
            Else
                dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52) / 365
            End If
        End If
        
        'Ticket #17654 - formula correction
        dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * Val(fglbDhrs)
    End If
    
ElseIf Data1.Recordset("JB_SALCD") = "A" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        If dblHoursPerWeek# = 0 Then
            dblNewSalary = -1
        Else
            dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52)
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) / 12
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##"))
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * 366
        Else
            dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * 365
        End If
        
        'Ticket #17654 - formula correction
        dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52) * Val(fglbDhrs)
    End If
End If
End Sub
Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbFrench Then
    tmpNUM = Replace(Replace(tmpNUM, ",", "."), " ", "")
    tmpNUM = Val(tmpNUM)
End If

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function

Private Sub modSetCOMPA_GRADE(dblNewSalary)

Dim X%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset
'SET COMPA RATIO
'================
SQLQ = "SELECT * FROM HRJOB"

snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
ssalary@ = dblNewSalary
dblHoursPerWeek# = dynSH_Job1("SH_WHRS")

'D added by Bryan 28/Sep/05 Ticket#9354
If Data1.Recordset("JB_SALCD") = "H" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        dblSsalary# = dblNewSalary
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary * 12) / (dblHoursPerWeek# * 52)
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = dblNewSalary / (dblHoursPerWeek# * 52)
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = dblNewSalary * 366 / (dblHoursPerWeek# * 52)
            Else
                dblSsalary# = dblNewSalary * 365 / (dblHoursPerWeek# * 52)
            End If
        
            'Ticket #17654 - formula correction
            dblSsalary# = dblNewSalary / Val(fglbDhrs)
        End If
    End If
ElseIf Data1.Recordset("JB_SALCD") = "A" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        dblSsalary# = dblNewSalary
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = dblNewSalary * 366
        Else
            dblSsalary# = dblNewSalary * 365
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = (dblNewSalary / Val(fglbDhrs)) * (dblHoursPerWeek# * 52)
    End If
End If

 ' set COMPA RATIO
 'laura 03/23/98

'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 11 Then
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 15 Then
If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 20 Then
    Jb_No = dynSH_Job1("JB_S" & dynSH_Job1("JB_MIDPOINT"))
End If

fglbCOMPA# = 0

If Jb_No <> 0 And dblSsalary# <> 0 Then 'laura 03/23/98
  fglbCOMPA# = (dblSsalary# / Jb_No) * 100
End If

 
If fglbCOMPA# > 999.99 Then
    fglbCOMPA# = 999.99
End If


'Determine Pay Scale individual fits into
'==========================================
snapJob.Requery
snapJob.Find "JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
fglbGRADE$ = "00"
xSalGrade = dblNewSalary
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For X% = 1 To 11
'For X% = 1 To 15
For X% = 1 To 20
    If IsNumeric(dynSH_Job1("JB_S" & Format(X%, "##"))) Then
        If snapJob("JB_SALCD") = "H" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                xSalGrade = snapJob("JB_S" & Format(X%, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                xSalGrade = (snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                xSalGrade = snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52) / 366
                Else
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52) / 365
                End If
                
                'Ticket #17654 - formula correction
                xSalGrade = snapJob("JB_S" & Format(X%, "##")) * Val(fglbDhrs)
            End If
        ElseIf snapJob("JB_SALCD") = "A" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                If dblHoursPerWeek# = 0 Then
                    xSalGrade = 0
                Else
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52)
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                xSalGrade = snapJob("JB_S" & Format(X%, "##")) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                xSalGrade = snapJob("JB_S" & Format(X%, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 366
                Else
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 365
                End If
            
                'Ticket #17654 - formula correction
                xSalGrade = snapJob("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52) * Val(fglbDhrs)
            End If
        End If
        If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(X%, "##")) > 0 Then
            cX$ = CStr(X)
            If X% <= 9 Then cX$ = "0" & cX$
            fglbGRADE$ = cX$
        End If
    End If
Next X%

If IsNumeric(dynSH_Job1("JB_S1")) Then
    If dblSsalary# < dynSH_Job1("JB_S1") Then
        fglbGRADE$ = "00"
    End If
End If

End Sub

Private Function AUDITSALY()
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim TB As New ADODB.Recordset
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITSALY = False


TB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
TB.MoveFirst
TB.Find "ED_EMPNBR = " & empNo&
If Not TB.EOF Then
    xPT = ""
    If Not IsNull(TB("ED_PT")) Then xPT = TB("ED_PT")
    xDiv = ""
    If Not IsNull(TB("ED_DIV")) Then xDiv = TB("ED_DIV")
        
Else
    xPT = ""
    xDiv = ""
End If

'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'strFields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_JOB, AU_SALCD, "
strFields = strFields & "AU_WHRS, AU_SEDATE, AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_SREASON "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

If OSalary <> NSalary Then GoTo MODUPD
If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD
GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDiv

TA("AU_SALARY") = NSalary
TA("AU_OLDSAL") = OSalary
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_JOB") = OJOB1         ' FRANK 4/5/2000
TA("AU_SALCD") = OSalCD
TA("AU_WHRS") = dblWHours# 'ADDED BY RAUBREY 7/7/97
If OEDate <> NEDate Then TA("AU_SEDATE") = IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_SREASON") = glbGridReason

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = empNo&

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    TA("AU_LDATE") = Format(DateAdd("d", 14, NEDate), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        If CVDate(NEDate) > CVDate(Date) Then
            TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
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
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & empNo&
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
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function


Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
'Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

    If Len(Trim(glbGridNDate)) > 0 Then
        rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = empNo&
        rsTB("EF_FDATE") = CVDate(glbGridNDate)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(empNo&, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "SREV"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg

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

Function CheckGridStepChange() As Boolean
    Dim rsMain As New ADODB.Recordset
    Dim SQLQ As String
    Dim HighestStep As Byte, I As Byte
    Dim NotUpdate As String, WillUpdate As String, LongMsg As String
    
    On Error GoTo ErrorHandler
    
    SQLQ = "SELECT JH_EMPNBR "
    If glbOracle Then
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY, HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0"
    Else
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY INNER JOIN HR_SALARY_HISTORY "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
    End If
    SQLQ = SQLQ & " AND JH_JOB='" & txtPosition.Text & "' "
    SQLQ = SQLQ & " AND (JH_WHRS=0 OR JH_WHRS IS NULL) "
    
    If Left(comPayPer.Text, 1) = "H" Then
        SQLQ = SQLQ & "AND SH_SALCD='A'"
    Else
        SQLQ = SQLQ & "AND SH_SALCD='H'"
    End If
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do Until rsMain.EOF
        If InStr(NotUpdate, ShowEmpnbr(rsMain("JH_EMPNBR")) & " - No Hours") = 0 Then NotUpdate = NotUpdate & ShowEmpnbr(rsMain("JH_EMPNBR")) & " - No Hours per Week entered" & vbCrLf
        rsMain.MoveNext
    Loop
    rsMain.Close
    Set rsMain = Nothing
    
    SQLQ = "SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE (SH_GRADE='0' OR SH_GRADE='00' OR SH_GRADE IS NULL) "
    SQLQ = SQLQ & " AND SH_JOB='" & txtPosition.Text & "' AND SH_CURRENT<>0 "
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do Until rsMain.EOF
        NotUpdate = NotUpdate & ShowEmpnbr(rsMain("SH_EMPNBR")) & " - Previous salary was at Step 00" & vbCrLf
        rsMain.MoveNext
    Loop
    rsMain.Close
    Set rsMain = Nothing
    
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_JOB='" & txtPosition.Text
    SQLQ = SQLQ & "' AND JH_CURRENT<>0"
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do Until rsMain.EOF
        If InStr(NotUpdate, ShowEmpnbr(rsMain("JH_EMPNBR")) & " ") = 0 Then WillUpdate = WillUpdate & ShowEmpnbr(rsMain("JH_EMPNBR")) & vbCrLf
        rsMain.MoveNext
    Loop
    rsMain.Close
    Set rsMain = Nothing
    
    If Len(NotUpdate) > 0 Then
        Load frmMsgBox
        frmMsgBox.Caption = "Salary Update Confirmation"
        LongMsg = "The following employee salaries will not be updated.  Continue?" & vbCrLf
        LongMsg = LongMsg & NotUpdate & vbCrLf
        LongMsg = LongMsg & "The following employee salaries *will* be updated:" & vbCrLf
        LongMsg = LongMsg & WillUpdate & IIf(WillUpdate = "", "None", "")
        frmMsgBox.txtLongMsg = LongMsg
        
        'Hemu
        'If WillUpdate = "" Then frmMsgBox.cmdOK.Enabled = False
        If WillUpdate = "" Then frmMsgBox.cmdCancel.Enabled = False
        'Hemu
        
        frmMsgBox.Show 1

        If glbMsgBoxResult = vbOK And WillUpdate <> "" Then
            CheckGridStepChange = True
        Else
            CheckGridStepChange = False
        End If
    Else
        CheckGridStepChange = True
    End If
    Unload frmMsgBox
    Exit Function
    
ErrorHandler:
    Call ERR_Hndlr(Err.Number, Err.Description, "CheckGridStepChange", "HR_JOB_HISTORY", SQLQ)
    CheckGridStepChange = False
    Exit Function
    Resume
End Function


'Private Sub Display_Value()
'    Dim SQLQ
'    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'        Call Set_Control("B", Me)
'        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'        If glbtermopen Then
'            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'        Else
'            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        End If
'        Exit Sub
'    End If
'
'
'    SQLQ = "SELECT * FROM HRJOB where JB_ID=" & Data1.Recordset!JB_ID
'
'    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
'    lblJobID = rsDATA!JB_ID
'    Call Set_Control("R", Me, rsDATA)
'
'End Sub

Private Function modReportsTo(Job)
Dim SQLQ As String
Dim rsRptTo As New ADODB.Recordset
modReportsTo = "None"
rsRptTo.Open "SELECT JB_REPTAU,JB_CODE,JB_DESCR FROM HRJOB WHERE JB_REPTAU  = '" & Job & "'", gdbAdoIhr001, adOpenForwardOnly
If Not rsRptTo.EOF Then
   modReportsTo = rsRptTo("JB_CODE") & " - " & rsRptTo("JB_DESCR")
End If

End Function
Function mod_Upd_Pos_Totals(updPCtComp%)
Dim snapJobCount As New ADODB.Recordset
Dim rsHRJOB As New ADODB.Recordset
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
If Not glbWFC Then 'Ticket #25785 Franks 07/30/2014
    SQLQ = SQLQ & " HRJOB.JB_POINTS = 0, "
End If
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


rsHRJOB.Open "HRJOB", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    rsHRJOB.MoveFirst
    rsHRJOB.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJOB.EOF Then
        'rsHrJob.Edit
        rsHRJOB("JB_NBRFIL") = JobCount&
        rsHRJOB.Update
    End If

    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
        
    End If

    snapJobCount.MoveNext
Wend

snapJobCount.Close

If glbWFC Then  'Ticket #25785 Franks 07/30/2014
'Jerry - Woodbridge only:
'"   Change the Count Positions to read the terminated file too. We will be removing this logic later.
    SQLQ = "SELECT JH_JOB, COUNT(JH_EMPNBR) AS NoPosFilled FROM Term_JOB_HISTORY WHERE (JH_CURRENT <> 0) GROUP BY JH_JOB"
    If snapJobCount.State <> 0 Then snapJobCount.Close
    snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    While Not snapJobCount.BOF And Not snapJobCount.EOF
        Job$ = snapJobCount("JH_JOB")
        JobCount& = snapJobCount("NoPosFilled")
        rsHRJOB.MoveFirst
        rsHRJOB.Find "JB_CODE = '" & Job$ & "'"
        
        If Not rsHRJOB.EOF Then
            'rsHrJob.Edit
            rsHRJOB("JB_NBRFIL") = rsHRJOB("JB_NBRFIL") + JobCount& 'added term jobs
            rsHRJOB.Update
        End If
    
        ''pct# = pct# + ipct#
        ''spct% = CInt(pct#)
        ''
        ''If updPCtComp Then
        ''    If pct# > 100 Then pct# = 0
        ''    MDIMain.panHelp(0).FloodPercent = pct#
        ''
        ''End If
    
        snapJobCount.MoveNext
    Wend
    
    snapJobCount.Close
    
    'Ticket #25969 - Franks 09/09/2014 - begin
    SQLQ = "SELECT * FROM HRJOB WHERE JB_NBRFIL = 0 "
    If snapJobCount.State <> 0 Then snapJobCount.Close
    snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not snapJobCount.EOF
        snapJobCount("JB_STATUS") = "INAC"
        If Not Left(snapJobCount("JB_DESCR"), 2) = "Z " Then
            snapJobCount("JB_DESCR") = Left(("Z DO NOT USE " & snapJobCount("JB_DESCR")), 50)
        End If
        snapJobCount.Update
        snapJobCount.MoveNext
    Loop
    snapJobCount.Close
    'Ticket #25969 - Franks 09/09/2014 - end
    
End If 'HRJOB.JB_NBRFIL = 0

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
    rsHRJOB.MoveFirst
    rsHRJOB.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJOB.EOF Then
        'rsHrJob.Edit
        If Not glbWFC Then 'Ticket #25785 Franks 07/30/2014
            rsHRJOB("JB_POINTS") = JobPoints#
        End If
        rsHRJOB.Update
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
    rsHRJOB.MoveFirst
    rsHRJOB.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJOB.EOF Then
        'rsHrJob.Edit
        rsHRJOB("JB_FTETotNu") = FTENum#
        rsHRJOB.Update
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
    
    rsHRJOB.MoveFirst
    rsHRJOB.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJOB.EOF Then
        rsHRJOB("JB_FTETotHr") = FTEHrs&
        rsHRJOB.Update
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

rsHRJOB.Close


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

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePOS
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Job_Master
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
Call modSTUPD(TF)

End Sub
Private Sub Form_Deactivate()
glbUserUploadMode = SwitchForm: Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub



Public Sub Display_Value()
 Dim SQLQ
 If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
     Call Set_Control("B", Me)
     If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
     If glbtermopen Then
         rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
     Else
         rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
     End If
     lblJobID = 0
Else
     SQLQ = "SELECT * FROM HRJOB where JB_ID=" & Data1.Recordset!JB_ID
 
     If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
     rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
 
     If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
     lblJobID = rsDATA!JB_ID
     Call Set_Control("R", Me, rsDATA)
End If

Call SET_UP_MODE
Call INI_GridStep
If oJobCode <> txtPosition Then
    Dim xForm As Form
    For Each xForm In Forms
        If xForm.name <> "MDIMain" And xForm.name <> Me.name Then
            If xForm.MDIChild Then
                If get_RelateMode(xForm) = RelatePOS Then
                    glbUserUploadMode = UploadFormWithoutCheck
                    Unload xForm
                End If
            End If
        End If
    Next
End If
oJobCode = txtPosition
oJobDesc = txtPosDescr
oJobUnion = clpCode(3).Text


'Ticket #28340 Franks 03/21/2016
lblMissingBudPos.Visible = IsMissingBudPos(txtPosition.Text)

End Sub


Public Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

'Data1.RecordSource = "SELECT * FROM HRJOB ORDER BY JB_DESCR"
SQLQ = SQLQ & "SELECT * FROM HRJOB WHERE 1 = 1"
If Len(glbWFCUserSecList) > 0 Then 'Ticket #27609 Franks 10/13/2015
    SQLQ = SQLQ & " AND JB_SECTION IN " & glbWFCUserSecList & " "
End If
If Len(clpCode(12).Text) > 0 Then 'Ticket #29552 Franks 12/14/2016
    SQLQ = SQLQ & "AND JB_SECTION = '" & clpCode(12).Text & "' "
End If
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    If glbOracle Then 'Ticket #16416
        SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
    Else
        SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
    End If
End If
SQLQ = SQLQ & "ORDER BY JB_DESCR"

Data1.RecordSource = SQLQ '"SELECT * FROM HRJOB ORDER BY JB_DESCR"
Data1.Refresh

If glbPos <> "" Then
    Data1.Recordset.Find "JB_CODE='" & glbPos & "' "
    If Data1.Recordset.RecordCount > 7 Then
        vbxTrueGrid.ScrollBars = vbVertical
    End If
End If

EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HRJOB", "SELECT")
Resume Next

Exit Function

End Function

Private Sub INI_GridStep()
Dim X

If SkipResetGridStep Then Exit Sub

If fglbNew Then
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        GridStep(X, 0) = 0
        GridStep(X, 2) = "N"
        IfGridStepChange = False

        'Hemu - Town of Aurora only - Ticket #10263
        If glbCompSerial = "S/N - 2378W" Then
            GridStep2(X, 0) = 0
            GridStep2(X, 2) = "N"
        End If
    Next X
    Exit Sub
End If
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        GridStep(X, 0) = 0
        GridStep(X, 2) = "N"
        IfGridStepChange = False
        
        'Hemu - Town of Aurora only - Ticket #10263
        If glbCompSerial = "S/N - 2378W" Then
            GridStep2(X, 0) = 0
            GridStep2(X, 2) = "N"
        End If
    Next X
    Exit Sub
End If
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For X = 1 To 11
'For X = 1 To 15
For X = 1 To 20
    If IsNull(Data1.Recordset("JB_S" & LTrim(RTrim(Str(X))))) Then
        GridStep(X, 0) = 0
    Else
        GridStep(X, 0) = Data1.Recordset("JB_S" & LTrim(RTrim(Str(X))))
    End If
    GridStep(X, 2) = "N"
    IfGridStepChange = False
    
    'Hemu - Town of Aurora only - Ticket #10263
    If glbCompSerial = "S/N - 2378W" Then
        If IsNull(Data1.Recordset("JB_S" & LTrim(RTrim(Str(X))) & "A")) Then
            GridStep2(X, 0) = 0
        Else
            GridStep2(X, 0) = Data1.Recordset("JB_S" & LTrim(RTrim(Str(X)) & "A"))
        End If
        GridStep2(X, 2) = "N"
    End If
Next X
End Sub


Private Sub Update_CompRatio()
Dim fTablSalHis As New ADODB.Recordset
Dim rsSH_Job1 As New ADODB.Recordset
Dim Emp_List As New Collection
Dim lngLastCurrentID&
Dim Jb_No#
Dim num
Dim dblSalary
Dim SQLQ
Dim tmpCOMPA
    
    Screen.MousePointer = HOURGLASS
    
    fTablSalHis.Open "HR_SALARY_HISTORY", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If glbOracle Then
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY,HRJOB WHERE HRJOB.JB_CODE = HR_SALARY_HISTORY.SH_JOB "
        SQLQ = SQLQ & " AND SH_CURRENT <> 0 and SH_JOB = '" & txtPosition & "'"
    Else
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_SALARY_HISTORY.SH_JOB "
        SQLQ = SQLQ & " WHERE SH_CURRENT <> 0 and SH_JOB = '" & txtPosition & "'"
    End If
    rsSH_Job1.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsSH_Job1.EOF Then
        rsSH_Job1.MoveFirst
        Do While Not rsSH_Job1.EOF
            Emp_List.Add (rsSH_Job1("SH_ID"))
            rsSH_Job1.MoveNext
        Loop
        
        DoEvents
        
        For num = 1 To Emp_List.count
            rsSH_Job1.MoveFirst
            lngLastCurrentID& = Emp_List(num)
            rsSH_Job1.Find "SH_ID = " & lngLastCurrentID&
            DoEvents

            fTablSalHis.MoveFirst
            fTablSalHis.Find "SH_ID = " & lngLastCurrentID&
            If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
                dblSalary = fTablSalHis("SH_SALARY") + fTablSalHis("SH_PREMIUM")
            Else
                dblSalary = fTablSalHis("SH_SALARY")
            End If

            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'If rsSH_Job1("JB_MIDPOINT") >= 1 And rsSH_Job1("JB_MIDPOINT") <= 11 Then
            'If rsSH_Job1("JB_MIDPOINT") >= 1 And rsSH_Job1("JB_MIDPOINT") <= 15 Then
            If rsSH_Job1("JB_MIDPOINT") >= 1 And rsSH_Job1("JB_MIDPOINT") <= 20 Then
                Jb_No = rsSH_Job1("JB_S" & rsSH_Job1("JB_MIDPOINT"))
            End If
            
            tmpCOMPA = 0
            
            If Jb_No <> 0 And dblSalary <> 0 Then
              tmpCOMPA = (dblSalary / Jb_No) * 100
            End If
             
            If tmpCOMPA > 999.99 Then
                tmpCOMPA = 999.99
            End If
            
            fTablSalHis("SH_COMPA") = Round(tmpCOMPA, 2)
            fTablSalHis("SH_LDATE") = Now
            fTablSalHis("SH_LTIME") = Time$
            fTablSalHis("SH_LUSER") = glbUserID
            fTablSalHis.Update
        Next
    End If
    rsSH_Job1.Close
    fTablSalHis.Close
    
    Screen.MousePointer = DEFAULT
    
End Sub

Private Sub Transfer_Salary(rsNew As ADODB.Recordset)
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim HRChanges As New Collection
    Dim UptSalaryDate As Date
    Dim HRSalary As New Collection
    Dim xEmpnbr
    Dim xPayrollID
    Dim xPHrs
    Dim xWhrs, xNiagaraWHRS
    Dim xEDate
    Dim xSalCD
    Dim UpdateAudit
    
    xEmpnbr = rsNew("SH_EMPNBR")
    If rsNew("SH_PAYROLL_ID") = "" Or IsNull(rsNew("SH_PAYROLL_ID")) Then
        xPayrollID = GetEmpData(rsNew("SH_EMPNBR"), "ED_PAYROLL_ID")
    Else
        xPayrollID = rsNew("SH_PAYROLL_ID")
    End If
    xEDate = rsNew("SH_EDATE")
    
    rsEmpJob.Open "SELECT JH_ID,JH_JOB,JH_DHRS,JH_PHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpnbr & " AND JH_PAYROLL_ID='" & xPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
    xPHrs = 0
    xWhrs = 0
    If Not rsEmpJob.EOF Then
        xPHrs = Val(rsEmpJob("JH_PHRS") & "")
        xWhrs = Val(rsEmpJob("JH_WHRS") & "") 'Hemu - it was asssigning JH_DHRS - it should pass Weekly Hours
        xNiagaraWHRS = Val(rsEmpJob("JH_WHRS") & "")
        
        'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
        If glbCompSerial = "S/N - 2276W" Then
            rsSal.Open "SELECT SH_EMPNBR, SH_PAYP, SH_WHRS FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpnbr & " AND SH_PAYROLL_ID = '" & xPayrollID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                xPHrs = Val(rsSal("SH_PAYP") & "")
                xNiagaraWHRS = Val(rsSal("SH_WHRS") & "")
            End If
            rsSal.Close
            Set rsSal = Nothing
            xWhrs = GetJobData(rsEmpJob("JH_JOB"), "JB_DHRS", 1)
            xWhrs = Val(xWhrs & "")
        End If
    End If
    rsEmpJob.Close
   
    If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
        If isChanged_Salary(HRSalary, OTOTAL, rsNew("SH_TOTAL"), True) Then UpdateAudit = True
    Else
        If isChanged_Salary(HRSalary, OSalary, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'Ticket #21352 - City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, Date, xPHrs, xWhrs, xEmpnbr, xPayrollID, , xNiagaraWHRS)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWhrs, xEmpnbr, xPayrollID, , xNiagaraWHRS)
        End If
    End If
    
    'Ticket #24565 - District Municipality of Muskoka
    If glbCompSerial = "S/N - 2373W" Then
        'They want to transfer for 181W as well now - Nov 3rd 2014
        'Ticket #24565 - if Union = '181W' then do not transfer Probation Date, Level and After Probation
        'If GetEmpData(xEmpnbr, "ED_ORG") = "181W" Then
        '    'Do not transfer Probation Date, Level and After Probation
        'Else
            If isChanged_Field(HRChanges, oGrade, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        'End If
    Else
        'Ticket #25469 - City of Campbell River - do not transfer Probation levels
        If glbCompSerial <> "S/N - 2458W" Then
            If isChanged_Field(HRChanges, oGrade, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        End If
    End If
    
    If isChanged_Field(HRChanges, OEDate, rsNew("SH_EDATE")) Then UpdateAudit = True
    If glbCompSerial <> "S/N - 2373W" Then 'DMuskoka - Ticket #24565 - Do not transfer Next Review Date
        If isChanged_Field(HRChanges, ONDate, rsNew("SH_NEXTDAT")) Then UpdateAudit = True
    End If
    Call Passing_Changes(HRChanges, Salary, "M", Date, xEmpnbr, xPayrollID)

End Sub

Private Sub Transfer_Position_Master_Vadim()
Dim xOccCode
Dim UpdType
Dim SQLQ
Dim pct#, prec%
Dim totalrec
If Not glbVadim Then Exit Sub


If fglbNew Then
    UpdType = "A"
    xOccCode = txtPosition
    If glbLambton Then xOccCode = Left(Left(clpGrid, 1) & txtPosition & Mid(clpGrid, 2), 6)
    Call Passing_Position_Master_Vadim(xOccCode, UpdType, oJobDesc, txtPosDescr)
Else
    UpdType = "M"
    If glbLambton Then
        Dim rsGridCategory As New ADODB.Recordset
        SQLQ = "SELECT JB_CODE,JB_GRID FROM HRJOB_GRADE WHERE JB_CODE='" & txtPosition & "'"
        rsGridCategory.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Do Until rsGridCategory.EOF
            xOccCode = Left(Left(rsGridCategory("JB_GRID") & "", 1) & txtPosition & Mid(rsGridCategory("JB_GRID") & "", 2), 6)
            Call Passing_Position_Master_Vadim(xOccCode, UpdType, oJobDesc, txtPosDescr)
            rsGridCategory.MoveNext
        Loop
        rsGridCategory.Close
    Else
        xOccCode = txtPosition
        Call Passing_Position_Master_Vadim(xOccCode, UpdType, oJobDesc, txtPosDescr)
    End If
End If

'Changes for the employees
If glbCompSerial = "S/N - 2363W" Then   ''City of Kawartha Lakes
    If (oJobDesc = txtPosDescr) And (oJobUnion = clpCode(3).Text) Then Exit Sub
Else
    If oJobDesc = txtPosDescr Then Exit Sub
End If
Dim rsEmpJob As New ADODB.Recordset
SQLQ = "SELECT JH_JOB,JH_PAYROLL_ID,JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB='" & txtPosition & "'"
SQLQ = SQLQ & " UNION "
SQLQ = SQLQ & "SELECT JH_JOB,JH_PAYROLL_ID,JH_EMPNBR FROM Term_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB='" & txtPosition & "'"
rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
totalrec = rsEmpJob.RecordCount
prec% = 0
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodShowPct = True
Do Until rsEmpJob.EOF
    prec% = prec% + 1
    MDIMain.panHelp(0).FloodPercent = (prec% / totalrec) * 100
    
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
        If (oJobUnion <> clpCode(3).Text) Then
            Call Transfer_Employee_Position_Vadim(rsEmpJob("JH_EMPNBR"), oJobUnion, clpCode(3), rsEmpJob("JH_PAYROLL_ID"))
        End If
        If (oJobDesc <> txtPosDescr) Then
            Call Transfer_Employee_Position_Vadim(rsEmpJob("JH_EMPNBR"), oJobDesc, txtPosDescr, rsEmpJob("JH_PAYROLL_ID"))
        End If
    Else
        Call Transfer_Employee_Position_Vadim(rsEmpJob("JH_EMPNBR"), oJobDesc, txtPosDescr, rsEmpJob("JH_PAYROLL_ID"))
    End If
    rsEmpJob.MoveNext
    DoEvents
Loop
rsEmpJob.Close
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodShowPct = False

End Sub

Private Sub Transfer_Employee_Position_Vadim(xEmpnbr, oldValue, InField As Control, xPayID)
Dim HRChanges As New Collection

If Not glbVadim Then Exit Sub

Call isChanged_Field(HRChanges, oldValue, InField)
Call Passing_Changes(HRChanges, PositionMaster, "M", Date, xEmpnbr, xPayID)

End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmMPOSITIONS")
    Call FillMemoFile(SQLQ, "Jobdescription")
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Jobdescription"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmMPOSITIONS")
End Sub


Private Function getWSQLQRPT(strEMPLIST As String) As String
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

Private Function getRows(exSheet As Object)
Dim X
X = 1
Do While True
    If exSheet.Cells(X, 1) = "" Then
        Exit Do
    Else
        X = X + 1
    End If
Loop
getRows = X - 1
End Function

Private Sub WFCOldPosUpt(xNewCode, xOldCode, xEndDate)
Dim SQLQ As String
    If xNewCode = xOldCode Then
        Exit Sub
    End If
    '"   Change Old Position Status to INAC and update the End Date
    SQLQ = "UPDATE HRJOB SET JB_STATUS = 'INAC' WHERE JB_CODE = '" & xOldCode & "' AND NOT JB_STATUS = 'INAC'"
    gdbAdoIhr001.Execute SQLQ
    If IsDate(xEndDate) Then
        SQLQ = "UPDATE HRJOB SET JB_EDATE = " & Date_SQL(xEndDate) & " WHERE JB_CODE = '" & xOldCode & "' AND NOT JB_STATUS = 'INAC'"
        gdbAdoIhr001.Execute SQLQ
    End If
    '"   Find all Positions that have their Reporting Authority 1 = the old Position Code and change the RA1 to equal the new Position Code.
    'Ignore any position with Status <>INAC .
    SQLQ = "UPDATE HRJOB SET JB_REPTAU = '" & xNewCode & "' WHERE JB_REPTAU = '" & xOldCode & "' AND NOT JB_STATUS = 'INAC'"
    gdbAdoIhr001.Execute SQLQ
    
    'Update Budgeted Position - begin
    SQLQ = "UPDATE HRJOBBUD SET JG_BUDGNBR = 0 WHERE JG_CODE = '" & xOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRJOBBUD SET JG_FTENUM = 0 WHERE JG_CODE = '" & xOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    Call mod_Upd_Pos_Budget_WFC(xOldCode, "")

    'Update Budgeted Position - end
End Sub

Private Sub WFCReptChaUptPopUp()
Dim rsEListWRK As New ADODB.Recordset
Dim xEmpNo, xPlant
Dim SQLQ As String

    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEListWRK.EOF Then
        'glbWFC_IPPopFormName = "WFCEmpListWithRept"
        glbWFC_IPPopFormName = "WFCEmpListWithRepPosBased"
        'glbWFC_IncePlanID = glbLEE_ID 'Employee based
        'glbWFC_IncePlanID = -100 'Position Master based
        frmCheckListView.lblStDate = dlpSDATE.Text
        frmCheckListView.Show 1
    End If
    rsEListWRK.Close
    
End Sub


Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'dlpDate(2).Text = Updstats(0)
        lblUptDate.Caption = Updstats(0)
    End If
    
    If Index = 2 Then
        lblUserDesc = GetUserDesc(Updstats(2))
    End If
End Sub
