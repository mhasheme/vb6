VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMPOSITIONS 
   Appearance      =   0  'Flat
   Caption         =   "Positions Master"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   960
   ClientWidth     =   11895
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
   ScaleWidth      =   11895
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
      TabIndex        =   67
      Top             =   9840
      Width           =   11895
      _Version        =   65536
      _ExtentX        =   20981
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
      Begin VB.CommandButton cmdCopyPosition 
         Appearance      =   0  'Flat
         Caption         =   "Copy Position..."
         Height          =   495
         Left            =   2520
         TabIndex        =   69
         Tag             =   "Copy selected Position info. to a New Position"
         Top             =   0
         Width           =   1905
      End
      Begin VB.CommandButton cmdAttachJobFiles 
         Appearance      =   0  'Flat
         Caption         =   "&Job Files..."
         Height          =   495
         Left            =   4800
         TabIndex        =   70
         Tag             =   "Attach Files related to this Job"
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.HScrollBar scrHScroll 
         Height          =   300
         LargeChange     =   25
         Left            =   0
         Max             =   50
         SmallChange     =   4
         TabIndex        =   136
         Top             =   520
         Width           =   11535
      End
      Begin VB.CommandButton cmdCountPos 
         Appearance      =   0  'Flat
         Caption         =   "&Count Positions + Total Points"
         Height          =   495
         Left            =   240
         TabIndex        =   68
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   0
         Width           =   1905
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9105
         Top             =   0
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
      TabIndex        =   71
      Top             =   2220
      Width           =   300
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmpostn.frx":0000
      Height          =   1815
      Left            =   30
      OleObjectBlob   =   "fxmpostn.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   11475
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   0
      TabIndex        =   76
      Top             =   1860
      Width           =   11475
      Begin VB.Frame frmVitalAireJobFamily 
         Height          =   1050
         Left            =   0
         TabIndex        =   158
         Top             =   6240
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   0
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "00-Bonus Reporting #"
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   1
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "00-Bonus Reporting #"
            Top             =   330
            Width           =   975
         End
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   2
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "00-Bonus Reporting #"
            Top             =   660
            Width           =   975
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   1305
            Picture         =   "fxmpostn.frx":11618
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblDouDivDesc 
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
            Index           =   0
            Left            =   2640
            TabIndex        =   164
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Family"
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
            Left            =   30
            TabIndex        =   163
            Top             =   0
            Width           =   735
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   1305
            Picture         =   "fxmpostn.frx":11762
            Top             =   330
            Width           =   240
         End
         Begin VB.Label lblDouDivDesc 
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
            Index           =   1
            Left            =   2640
            TabIndex        =   162
            Top             =   330
            Width           =   4095
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Job Family"
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
            Left            =   30
            TabIndex        =   161
            Top             =   350
            Width           =   1065
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   1305
            Picture         =   "fxmpostn.frx":118AC
            Top             =   660
            Width           =   240
         End
         Begin VB.Label lblDouDivDesc 
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
            Index           =   2
            Left            =   2640
            TabIndex        =   160
            Top             =   660
            Width           =   4095
         End
         Begin VB.Label lblJobF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Group Jobs"
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
            Left            =   30
            TabIndex        =   159
            Top             =   680
            Width           =   810
         End
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
         TabIndex        =   139
         Top             =   98
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
         Left            =   9120
         TabIndex        =   123
         Top             =   480
         Width           =   2250
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S10A"
            Height          =   285
            Index           =   30
            Left            =   480
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   56
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
            TabIndex        =   57
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
            TabIndex        =   58
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
            TabIndex        =   59
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
            TabIndex        =   60
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
            TabIndex        =   61
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
            TabIndex        =   62
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
            TabIndex        =   63
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
            TabIndex        =   64
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
            TabIndex        =   65
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
            TabIndex        =   157
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
            TabIndex        =   156
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            TabIndex        =   147
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
            TabIndex        =   146
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   134
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
            TabIndex        =   133
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
            Top             =   285
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   8100
         TabIndex        =   121
         Top             =   120
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
         Left            =   5280
         TabIndex        =   8
         Tag             =   "01-Mid Point Grid Step number"
         Text            =   "cmbMidPoint"
         Top             =   1320
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
         Left            =   5670
         MaxLength       =   2
         TabIndex        =   89
         Top             =   1680
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
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   14
         Tag             =   "10-Number of positions that exist for this job"
         Top             =   3240
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
         Left            =   6720
         TabIndex        =   77
         Top             =   480
         Width           =   2250
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S10"
            Height          =   285
            Index           =   10
            Left            =   480
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   36
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   152
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
            TabIndex        =   151
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            TabIndex        =   148
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
            Top             =   285
            Width           =   90
         End
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         DataField       =   "JB_CODE"
         Height          =   285
         Left            =   780
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "01-Position Code (Unique)"
         Top             =   30
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
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "01-Position Description"
         Top             =   30
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "00-Position Alternate Description"
         Top             =   360
         Width           =   4030
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
         TabIndex        =   15
         Tag             =   "10-Number of FTE "
         Top             =   3540
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
         TabIndex        =   16
         Tag             =   "10-FTE Hours/Year"
         Top             =   3840
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
         TabIndex        =   6
         Tag             =   "01- Grid Steps - Annual or Hourly"
         Top             =   1350
         Width           =   2730
      End
      Begin INFOHR_Controls.CodeLookup clpLGroup 
         DataField       =   "JB_LOCGROUP"
         Height          =   285
         Left            =   1305
         TabIndex        =   21
         Tag             =   "Location Group"
         Top             =   5490
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
         Left            =   1300
         TabIndex        =   11
         Tag             =   "00-Enter Position Code"
         Top             =   2310
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU2"
         Height          =   285
         Index           =   1
         Left            =   1300
         TabIndex        =   10
         Tag             =   "00-Enter Position Code"
         Top             =   2000
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpReportsTo 
         DataField       =   "JB_REPTAU"
         Height          =   285
         Index           =   0
         Left            =   1300
         TabIndex        =   9
         Tag             =   "00-Enter Position Code"
         Top             =   1680
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpNationalClass 
         DataField       =   "JB_FEDGRP"
         Height          =   285
         Left            =   1310
         TabIndex        =   13
         Tag             =   "00-National Occupation Classification -Code"
         Top             =   2940
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
         Left            =   1305
         TabIndex        =   20
         Tag             =   "01-WF2 Code"
         Top             =   5160
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
         Left            =   1305
         TabIndex        =   19
         Tag             =   "01-WF1 Code"
         Top             =   4830
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
         Left            =   1300
         TabIndex        =   12
         Tag             =   "00-Union - Code "
         Top             =   2620
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_GRPCD"
         Height          =   285
         Index           =   2
         Left            =   1300
         TabIndex        =   5
         Tag             =   "01-Position Group Code "
         Top             =   1020
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_STATUS"
         Height          =   285
         Index           =   1
         Left            =   1300
         TabIndex        =   4
         Tag             =   "01-Position Status - Code "
         Top             =   690
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBST"
         MaxLength       =   6
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JB_BAND"
         Height          =   285
         Index           =   6
         Left            =   3720
         TabIndex        =   7
         Tag             =   "00-Band - Code"
         Top             =   1350
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   5820
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
         TabIndex        =   18
         Tag             =   "10-Usual working hours per day"
         Top             =   4500
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
         Left            =   1305
         TabIndex        =   137
         Tag             =   "00-Machine #"
         Top             =   4830
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
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Tag             =   "10-Points"
         Top             =   4170
         Visible         =   0   'False
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
         Left            =   30
         TabIndex        =   138
         Top             =   4890
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
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
         TabIndex        =   135
         Top             =   4545
         Width           =   930
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   7680
         Picture         =   "fxmpostn.frx":119F6
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   7680
         Picture         =   "fxmpostn.frx":11B40
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5880
         TabIndex        =   122
         Top             =   120
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
         TabIndex        =   120
         Top             =   7380
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Occupation"
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
         Left            =   6330
         TabIndex        =   119
         Top             =   7410
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblGridC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Category (Default)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   118
         Top             =   5880
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label lblJobID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4980
         TabIndex        =   117
         Top             =   780
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   116
         Top             =   735
         Width           =   555
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   115
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label lblSalary 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Grid"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   114
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 1"
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
         Left            =   30
         TabIndex        =   113
         Top             =   1695
         Width           =   930
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
         Left            =   30
         TabIndex        =   112
         Top             =   2610
         Width           =   420
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
         Left            =   30
         TabIndex        =   111
         Top             =   2925
         Width           =   900
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
         Left            =   30
         TabIndex        =   110
         Top             =   3240
         Width           =   1410
      End
      Begin VB.Label lblFTENum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE #"
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
         TabIndex        =   109
         Top             =   3555
         Width           =   540
      End
      Begin VB.Label lblFTEHrs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE Hours/Year"
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
         TabIndex        =   108
         Top             =   3870
         Width           =   1395
      End
      Begin VB.Label lblTotalPoints 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Points"
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
         TabIndex        =   107
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_POINTS"
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
         Left            =   1620
         TabIndex        =   106
         Top             =   4190
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
         Left            =   3900
         TabIndex        =   105
         Top             =   1425
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
         Left            =   3285
         TabIndex        =   104
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label lblTotNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE #"
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
         Left            =   3285
         TabIndex        =   103
         Top             =   3555
         Width           =   1035
      End
      Begin VB.Label lblTotHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE Hours/Year"
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
         Left            =   3285
         TabIndex        =   102
         Top             =   3900
         Width           =   1575
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
         Left            =   3330
         TabIndex        =   101
         Top             =   4150
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
         Left            =   5100
         TabIndex        =   100
         Top             =   3225
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
         TabIndex        =   99
         Top             =   3555
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
         Left            =   5100
         TabIndex        =   98
         Top             =   3900
         Width           =   90
      End
      Begin VB.Label lblComRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mid-Point"
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
         Left            =   4560
         TabIndex        =   97
         Top             =   1365
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
         TabIndex        =   66
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate"
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
         Left            =   165
         TabIndex        =   96
         Top             =   360
         Width           =   1350
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
         TabIndex        =   95
         Top             =   1380
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
         Left            =   30
         TabIndex        =   94
         Top             =   2310
         Width           =   930
      End
      Begin VB.Label lblReptAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports To 2"
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
         Left            =   30
         TabIndex        =   93
         Top             =   2010
         Width           =   930
      End
      Begin VB.Label lblWF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "WF1"
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
         Left            =   30
         TabIndex        =   92
         Top             =   4890
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblWF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "WF2"
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
         Left            =   30
         TabIndex        =   91
         Top             =   5220
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location Group "
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
         TabIndex        =   90
         Top             =   5520
         Visible         =   0   'False
         Width           =   1140
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
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   1500
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
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   74
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   1500
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
      Left            =   720
      MaxLength       =   25
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   1500
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
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmMPOSITIONS"
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
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, EmpNo&, dblWHours#, OTotal
Dim oPayP, NPayp, OJOB1, OSalCD, oGrade
Dim EmpChgErrors, strEmpEffError As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim lstMidPoint
Dim fglbNew As Boolean
Dim oJobCode, oJobDesc, oJobUnion
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
If Not modJobUnique(Job, JobID) Then
    MsgBox "Job Code is not unique"
    txtPosition.Enabled = True
    txtPosition.SetFocus
    Exit Function
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
If glbCompSerial = "S/N - 2409W" Then 'Delisle Youth Services - Ticket #27798
    If Len(txtPosDescr2) = 0 Then
        MsgBox "Alternate is a required field"
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
For X = 0 To 2
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

Private Sub clpGrid_LostFocus()
    If txtPosition <> "" And clpGrid <> "" Then
        txtLambtonJob = Left(clpGrid, 1) & txtPosition & Mid(clpGrid, 2, 1)
    End If
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

Private Sub cmdCopyPosition_Click()
    'Call screen to enter a new Position Code.
    frmPosCopy.clpFromJob = txtPosition
    frmPosCopy.Show 1

    '------------------refreshing the form
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Call SET_UP_MODE
    Call Display_Value
    '----------------------

End Sub

Private Sub cmdCopyPosition_GotFocus()
Call SetPanHelp(ActiveControl)
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

On Error GoTo NewErr
'Call modSTUPD(False)

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
txtPosition.Enabled = True
txtPosition.SetFocus
Call INI_GridStep
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
On Error GoTo OK_Err

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
Dim RSTABL As New ADODB.Recordset
Dim X
rsGrid.Open "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & zJOB & "'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If clpGrid = "DFLT" Then
    RSTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='JBGD' AND TB_KEY='DFLT'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If RSTABL.EOF Then
        RSTABL.AddNew
        RSTABL("TB_NAME") = "JBGD"
        RSTABL("TB_KEY") = "DFLT"
        RSTABL("TB_DESC") = "DEFAULT"
        RSTABL.Update
    End If
    RSTABL.Close
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
    glbOnTop = "FRMMPOSITIONS"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMMPOSITIONS"
End Sub

Private Sub Form_Load()
flagLoad = 1
Dim SQLQ As String, RFound%, X As Integer

If Not glbWFC Then 'Ticket #20479 Franks 07/11/2011, this field is for WFC only
    vbxTrueGrid.Columns(5).Visible = False
End If
If glbSyndesis Then
    fraGrid.Caption = "Range"
    lblGroup.Caption = "Grade"
    vbxTrueGrid.Columns(4).Caption = "Grade"
    vbxTrueGrid.Columns(19).Caption = "Range 1"
    vbxTrueGrid.Columns(20).Caption = "Range 2"
    vbxTrueGrid.Columns(21).Caption = "Range 3"
    vbxTrueGrid.Columns(22).Caption = "Range 4"
    vbxTrueGrid.Columns(23).Caption = "Range 5"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15
    'vbxTrueGrid.Columns(24).Caption = "Range 6 (Mid)"
    vbxTrueGrid.Columns(24).Caption = "Range 6 "
    vbxTrueGrid.Columns(25).Caption = "Range 7"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15
    'vbxTrueGrid.Columns(26).Caption = "Range 8"
    vbxTrueGrid.Columns(26).Caption = "Range 8"
    vbxTrueGrid.Columns(27).Caption = "Range 9"
    vbxTrueGrid.Columns(28).Caption = "Range 10"
    vbxTrueGrid.Columns(29).Caption = "Range 11"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    vbxTrueGrid.Columns(30).Caption = "Range 12"
    vbxTrueGrid.Columns(31).Caption = "Range 13"
    vbxTrueGrid.Columns(32).Caption = "Range 14"
    vbxTrueGrid.Columns(33).Caption = "Range 15"
    
    vbxTrueGrid.Columns(34).Caption = "Range 16"
    vbxTrueGrid.Columns(35).Caption = "Range 17"
    vbxTrueGrid.Columns(36).Caption = "Range 18"
    vbxTrueGrid.Columns(37).Caption = "Range 19"
    vbxTrueGrid.Columns(38).Caption = "Range 20"
    
ElseIf glbCompSerial = "S/N - 2366W" Then   'Family Youth and Child Services of Muskoka
    vbxTrueGrid.Columns(19).Caption = "Start"
    vbxTrueGrid.Columns(20).Caption = "Grid Step 1"
    vbxTrueGrid.Columns(21).Caption = "Grid Step 2"
    vbxTrueGrid.Columns(22).Caption = "Grid Step 3"
    vbxTrueGrid.Columns(23).Caption = "Grid Step 4"
    vbxTrueGrid.Columns(24).Caption = "Grid Step 5"
    vbxTrueGrid.Columns(25).Caption = "Grid Step 6"
    vbxTrueGrid.Columns(26).Caption = "Grid Step 7"
    vbxTrueGrid.Columns(27).Caption = "Grid Step 8"
    vbxTrueGrid.Columns(28).Caption = "Grid Step 9"
    vbxTrueGrid.Columns(29).Caption = "Grid Step 10"
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    vbxTrueGrid.Columns(30).Caption = "Grid Step 11"
    vbxTrueGrid.Columns(31).Caption = "Grid Step 12"
    vbxTrueGrid.Columns(32).Caption = "Grid Step 13"
    vbxTrueGrid.Columns(33).Caption = "Grid Step 14"
    
    vbxTrueGrid.Columns(34).Caption = "Grid Step 15"
    vbxTrueGrid.Columns(35).Caption = "Grid Step 16"
    vbxTrueGrid.Columns(36).Caption = "Grid Step 17"
    vbxTrueGrid.Columns(37).Caption = "Grid Step 18"
    vbxTrueGrid.Columns(38).Caption = "Grid Step 19"
    
ElseIf glbCompSerial = "S/N - 2380W" Then ' Vitalaire
    lblGroup.Caption = "Job Class"
    vbxTrueGrid.Columns(4).Caption = "Job Class"
    Call VitalAireJobFamilyScreen 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
ElseIf glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #17490
    Label1.Caption = "Ceridian Key"
    Label1.FontBold = True
    vbxTrueGrid.Columns(2).Caption = "Ceridian Key"
    txtPosDescr2.MaxLength = 2
    txtPosDescr2.Width = 800
ElseIf glbCompSerial = "S/N - 2172W" Then 'Lanark
    lblGroup.Caption = "Salary Level"
    vbxTrueGrid.Columns(4).Caption = "Salary Level"
    clpCode(2).TABLTitle = "Salary Level Code"
ElseIf glbCompSerial = "S/N - 2409W" Then 'Delisle Youth Services - Ticket #27798
    Label1.FontBold = True
End If

glbOnTop = "FRMMPOSITIONS"
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
If glbWFC Then clpCode(6).Left = 1300

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
Call setCaption(lblUnion)

lblStatus = Replace(lStr("Position " & lblStatus), "Position ", "")
lblGroup = Replace(lStr("Position " & lblGroup), "Position ", "")

vbxTrueGrid.Columns(3).Caption = Replace(lStr("Position Status"), "Position ", "")
vbxTrueGrid.Columns(4).Caption = Replace(lStr("Position Group"), "Position ", "")

If Not gSec_Upd_Job_Master Then        'May99 js
    Call set_Buttons
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call SetMultiGrid
Call INI_Controls(Me)
Call Display_Value

If glbWFC Then 'Ticket #25785 Franks 07/30/2014
    Call PointsTextboxScreenSetup
End If

Screen.MousePointer = DEFAULT
                             '
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
Set frmMPOSITIONS = Nothing   'carmen may 2000

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

Private Function modJobUnique(Job, JobID)
Dim SQLQ As String
Dim rsJob As New ADODB.Recordset
modJobUnique = True
rsJob.Open "SELECT JB_CODE FROM HRJOB WHERE JB_CODE='" & Trim(Job) & "' AND JB_ID<>" & JobID, gdbAdoIhr001, adOpenForwardOnly
If Not rsJob.EOF Then modJobUnique = False
rsJob.Close
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

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdCountPos.Enabled = False
    cmdCopyPosition.Enabled = False
Else
    If fglbNew Then
        cmdCountPos.Enabled = False
        cmdCopyPosition.Enabled = False
    Else
        cmdCountPos.Enabled = TF  'FT    '
        'Release 8.1
        cmdCopyPosition.Enabled = TF
    End If
End If

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

If glbCompSerial = "S/N - 2380W" Then 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
    txtDouDiv(0).Enabled = TF
    txtDouDiv(1).Enabled = TF
    txtDouDiv(2).Enabled = TF
End If

clpLGroup.Enabled = TF
End Sub

Private Sub medPayScale_LostFocus(Index As Integer)
    'Hemu - Ticket #10139 - Town of Aurora only
    'Oxford Ticket #15590
    If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2259W" Then
        Call Calculate_Secondary_Grid_Steps(Index)
    End If
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
                lblComRatio.Visible = True
                cmbMidPoint.Visible = True
                comPayPer.Visible = True
                lblSalary.Visible = True
                fraGrid.Visible = True
            End If
        End If
        comPayPer.Refresh
    End If
End Sub

Private Sub scrHScroll_Change()
fraDetail.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtDouDiv_Change(Index As Integer)
lblDouDivDesc(Index).Caption = getJobFamilyDesc(txtDouDiv(Index).Text, Index)
End Sub

Private Sub txtDouDiv_DblClick(Index As Integer)
    If Index = 0 Then
        Call Get_JobFamily(False, "JOBFAMILY")
        If Len(glbJobFam) > 0 Then
            txtDouDiv(0).Text = glbJobFam
        End If
    End If
    If Index = 1 Then
        Call Get_JobFamily(False, "SUBFAMILY", txtDouDiv(0).Text)
        If Len(glbSubJobFam) > 0 Then
            txtDouDiv(1).Text = glbSubJobFam
        End If
    End If
    If Index = 2 Then
        Call Get_JobFamily(False, "GROUPJOBS", txtDouDiv(1).Text)
        If Len(glbGroupJob) > 0 Then
            txtDouDiv(2).Text = glbGroupJob
        End If
    End If
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
            lblComRatio.Visible = True
            cmbMidPoint.Visible = True
            comPayPer.Visible = True
            lblSalary.Visible = True
            fraGrid.Visible = True
        End If
        comPayPer.Refresh
    End If
End If

If Data1.Recordset.EOF Then
    glbPos = ""
    glbPosDesc = ""
Else
    glbPos = Data1.Recordset("JB_CODE")
    glbPosDesc = Data1.Recordset("JB_DESCR")
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
On Error GoTo RR
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
RR:
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
        
        EmpNo& = dynSH_Job1("SH_EMPNBR")
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
        OTotal = fTablSalHis("SH_TOTAL")
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
        
        Call updBenefitForSalDEPN(EmpNo&)
        
        'Ticket #16991 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
        'remaining same, only the actual salary is changing and this table only stores the Rate Level
        'City of Niagara Falls - Ticket #15542
        'If glbVadim And glbCompSerial = "S/N - 2276W" Then
        '    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
        '    Call Update_VadimDB_HR_EMP_HISTORY(fTablSalHis("SH_PAYROLL_ID"), CVDate(glbGridEDate), "", Val(fglbGRADE$), fTablSalHis("SH_JOB"), "A")
        'End If
        
        'If Not glbWFC Then 'Greensboro
            Call Employee_Master_Integration(EmpNo&)
        'End If
        If glbGP Then Call Salary_Integration(EmpNo&, , False, True, xSHID)
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
TB.Find "ED_EMPNBR = " & EmpNo&
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
TA("AU_EMPNBR") = EmpNo&

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
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & EmpNo&
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
        rsTB("EF_EMPNBR") = EmpNo&
        rsTB("EF_FDATE") = CVDate(glbGridNDate)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(EmpNo&, "ED_ADMINBY", Null)
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

If glbWFC Then  'Ticket #25785 Franks 07/30/2014
'Jerry - Woodbridge only:
'"   Change the Count Positions to read the terminated file too. We will be removing this logic later.
    SQLQ = "SELECT JH_JOB, COUNT(JH_EMPNBR) AS NoPosFilled FROM Term_JOB_HISTORY WHERE (JH_CURRENT <> 0) GROUP BY JH_JOB"
    If snapJobCount.State <> 0 Then snapJobCount.Close
    snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    While Not snapJobCount.BOF And Not snapJobCount.EOF
        Job$ = snapJobCount("JH_JOB")
        JobCount& = snapJobCount("NoPosFilled")
        rsHRJob.MoveFirst
        rsHRJob.Find "JB_CODE = '" & Job$ & "'"
        
        If Not rsHRJob.EOF Then
            'rsHrJob.Edit
            rsHRJob("JB_NBRFIL") = rsHRJob("JB_NBRFIL") + JobCount& 'added term jobs
            rsHRJob.Update
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
    rsHRJob.MoveFirst
    rsHRJob.Find "JB_CODE = '" & Job$ & "'"
    
    If Not rsHRJob.EOF Then
        'rsHrJob.Edit
        If Not glbWFC Then 'Ticket #25785 Franks 07/30/2014
            rsHRJob("JB_POINTS") = JobPoints#
        End If
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

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePos
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
                If get_RelateMode(xForm) = RelatePos Then
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
End Sub


Public Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

'Data1.RecordSource = "SELECT * FROM HRJOB ORDER BY JB_DESCR"
SQLQ = SQLQ & "SELECT * FROM HRJOB WHERE 1 = 1"
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
    Dim xWHrs, xNiagaraWHRS
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
    xWHrs = 0
    If Not rsEmpJob.EOF Then
        xPHrs = Val(rsEmpJob("JH_PHRS") & "")
        xWHrs = Val(rsEmpJob("JH_WHRS") & "") 'Hemu - it was asssigning JH_DHRS - it should pass Weekly Hours
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
            xWHrs = GetJobData(rsEmpJob("JH_JOB"), "JB_DHRS", 1)
            xWHrs = Val(xWHrs & "")
        End If
    End If
    rsEmpJob.Close
   
    If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
        If isChanged_Salary(HRSalary, OTotal, rsNew("SH_TOTAL"), True) Then UpdateAudit = True
    Else
        If isChanged_Salary(HRSalary, OSalary, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'Ticket #21352 - City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, Date, xPHrs, xWHrs, xEmpnbr, xPayrollID, , xNiagaraWHRS)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWHrs, xEmpnbr, xPayrollID, , xNiagaraWHRS)
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
Private Sub VitalAireJobFamilyScreen() 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
    txtDouDiv(0).DataField = "JB_JOBFAMILY"
    txtDouDiv(1).DataField = "JB_SUBJOBFAMILY"
    txtDouDiv(2).DataField = "JB_JOBFAMILYGRP"
    frmVitalAireJobFamily.Left = 0
    frmVitalAireJobFamily.Top = 4830
    frmVitalAireJobFamily.BorderStyle = 0
    frmVitalAireJobFamily.Visible = True
End Sub


