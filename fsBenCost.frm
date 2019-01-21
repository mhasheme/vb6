VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSBenCost 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Benefit Costing Details"
   ClientHeight    =   7485
   ClientLeft      =   2565
   ClientTop       =   525
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7485
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame VacFram 
      BorderStyle     =   0  'None
      Height          =   3570
      Left            =   60
      TabIndex        =   83
      Top             =   2880
      Width           =   11000
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   0
         Left            =   5280
         TabIndex        =   104
         Top             =   -30
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Tag             =   "Entitlement measured in days"
            Top             =   150
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   1
         Left            =   5280
         TabIndex        =   95
         Top             =   300
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   14
            Tag             =   "Entitlement measured in days"
            Top             =   180
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   2
         Left            =   5280
         TabIndex        =   96
         Top             =   660
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   2
            Left            =   1140
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   21
            Tag             =   "Entitlement measured in days"
            Top             =   150
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   23
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   3
         Left            =   5280
         TabIndex        =   97
         Top             =   990
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   3
            Left            =   1140
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   28
            Tag             =   "Entitlement measured in days"
            Top             =   180
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   3
            Left            =   2040
            TabIndex        =   30
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   4
         Left            =   5280
         TabIndex        =   98
         Top             =   1350
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   4
            Left            =   1140
            TabIndex        =   36
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   35
            Tag             =   "Entitlement measured in days"
            Top             =   150
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   37
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   5
         Left            =   5280
         TabIndex        =   99
         Top             =   1670
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   5
            Left            =   1140
            TabIndex        =   43
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   42
            Tag             =   "Entitlement measured in days"
            Top             =   180
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   5
            Left            =   2040
            TabIndex        =   44
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   6
         Left            =   5280
         TabIndex        =   100
         Top             =   2020
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   6
            Left            =   1140
            TabIndex        =   50
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   49
            Tag             =   "Entitlement measured in days"
            Top             =   150
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   6
            Left            =   2040
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   435
         Index           =   7
         Left            =   5280
         TabIndex        =   101
         Top             =   2340
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   7
            Left            =   1140
            TabIndex        =   57
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   56
            Tag             =   "Entitlement measured in days"
            Top             =   180
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   7
            Left            =   2040
            TabIndex        =   58
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   180
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   430
         Index           =   8
         Left            =   5280
         TabIndex        =   102
         Top             =   2700
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   758
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   8
            Left            =   1140
            TabIndex        =   64
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   63
            Tag             =   "Entitlement measured in days"
            Top             =   150
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   8
            Left            =   2040
            TabIndex        =   65
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin Threed.SSFrame frmDH 
         Height          =   460
         Index           =   9
         Left            =   5280
         TabIndex        =   103
         Top             =   3000
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   811
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optM 
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   70
            Tag             =   "Entitlement measured in days"
            Top             =   210
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Monthly"
            ForeColor       =   -2147483640
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
         Begin Threed.SSOption optA 
            Height          =   195
            Index           =   9
            Left            =   1140
            TabIndex        =   71
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   210
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Annual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optW 
            Height          =   195
            Index           =   9
            Left            =   2040
            TabIndex        =   72
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   210
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Weekly"
            ForeColor       =   -2147483640
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
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   9
         Left            =   4380
         TabIndex        =   69
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   8
         Left            =   4380
         TabIndex        =   62
         Top             =   2790
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   7
         Left            =   4380
         TabIndex        =   55
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   6
         Left            =   4380
         TabIndex        =   48
         Top             =   2070
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   5
         Left            =   4380
         TabIndex        =   41
         Top             =   1740
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   4
         Left            =   4380
         TabIndex        =   34
         Top             =   1410
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   3
         Left            =   4380
         TabIndex        =   27
         Top             =   1050
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   2
         Left            =   4380
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   1
         Left            =   4380
         TabIndex        =   13
         Top             =   390
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkFTE 
         Caption         =   "FTE"
         Height          =   375
         Index           =   0
         Left            =   4380
         TabIndex        =   6
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   "11-Service is greater than this number"
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Tag             =   "10-Service is less than this number"
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Tag             =   "11-Service is greater than this number"
         Top             =   456
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Tag             =   "10-Service is less than this number"
         Top             =   456
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Tag             =   "11-Service is greater than this number"
         Top             =   792
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   18
         Tag             =   "10-Service is less than this number"
         Top             =   792
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   24
         Tag             =   "11-Service is greater than this number"
         Top             =   1128
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   25
         Tag             =   "10-Service is less than this number"
         Top             =   1128
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Tag             =   "11-Service is greater than this number"
         Top             =   1464
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   32
         Tag             =   "10-Service is less than this number"
         Top             =   1464
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   38
         Tag             =   "11-Service is greater than this number"
         Top             =   1800
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   39
         Tag             =   "10-Service is less than this number"
         Top             =   1800
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   45
         Tag             =   "11-Service is greater than this number"
         Top             =   2136
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   46
         Tag             =   "10-Service is less than this number"
         Top             =   2136
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   7
         Left            =   0
         TabIndex        =   52
         Tag             =   "11-Service is greater than this number"
         Top             =   2472
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   53
         Tag             =   "10-Service is less than this number"
         Top             =   2472
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   59
         Tag             =   "11-Service is greater than this number"
         Top             =   2808
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   60
         Tag             =   "10-Service is less than this number"
         Top             =   2808
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTSal 
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   66
         Tag             =   "11-Service is greater than this number"
         Top             =   3150
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTSal 
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   67
         Tag             =   "10-Service is less than this number"
         Top             =   3150
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "#0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Tag             =   "11-Entitlement Amount"
         Top             =   456
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   19
         Tag             =   "11-Entitlement Amount"
         Top             =   792
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   33
         Tag             =   "11-Entitlement Amount"
         Top             =   1464
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   40
         Tag             =   "11-Entitlement Amount"
         Top             =   1800
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   26
         Tag             =   "11-Entitlement Amount"
         Top             =   1128
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   54
         Tag             =   "11-Entitlement Amount"
         Top             =   2472
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   8
         Left            =   3240
         TabIndex        =   61
         Tag             =   "11-Entitlement Amount"
         Top             =   2808
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   47
         Tag             =   "11-Entitlement Amount"
         Top             =   2136
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   9
         Left            =   3240
         TabIndex        =   68
         Tag             =   "11-Entitlement Amount"
         Top             =   3150
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPC 
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   5
         Tag             =   "11-Entitlement Amount"
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
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
         Format          =   "##0.0000%"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Salary =>"
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
         Left            =   1050
         TabIndex        =   93
         Top             =   165
         Width           =   930
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Salary =>"
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
         Left            =   1050
         TabIndex        =   92
         Top             =   501
         Width           =   930
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   91
         Top             =   837
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   90
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   89
         Top             =   1509
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   88
         Top             =   1173
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   87
         Top             =   2853
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   86
         Top             =   2517
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<= Salary =>"
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
         Left            =   1065
         TabIndex        =   85
         Top             =   2181
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ">   Salary"
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
         Left            =   1080
         TabIndex        =   84
         Top             =   3195
         Width           =   675
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   3375
      LargeChange     =   315
      Left            =   10980
      Max             =   100
      SmallChange     =   315
      TabIndex        =   78
      Top             =   3210
      Width           =   300
   End
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   2925
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   11415
      Begin INFOHR_Controls.CodeLookup clpBGroup 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Tag             =   "00-Benefit Group"
         Top             =   1800
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsBenCost.frx":0000
         Height          =   1335
         Left            =   180
         OleObjectBlob   =   "fsBenCost.frx":0014
         TabIndex        =   0
         Top             =   0
         Width           =   9135
      End
      Begin INFOHR_Controls.CodeLookup clpBCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Tag             =   "00-Enter Benefit Code"
         Top             =   2130
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNCD"
         MaxLength       =   10
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3300
         TabIndex        =   94
         Top             =   2670
         Width           =   990
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   82
         Top             =   2670
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   81
         Top             =   2670
         Width           =   660
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   80
         Top             =   2940
         Width           =   75
      End
      Begin VB.Label lblBCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   79
         Top             =   2160
         Width           =   1110
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
         TabIndex        =   77
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblBGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Group"
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
         Left            =   180
         TabIndex        =   76
         Top             =   1830
         Width           =   975
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   73
      Top             =   6855
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1111
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   7080
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
End
Attribute VB_Name = "frmSBenCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbNew As Boolean
Dim OBGroup, OBCode
Dim fglbBSQLQ

Private Function chkMUBenCost()
Dim X%, Y%

chkMUBenCost = False

On Error GoTo chkMUBenCost_Err
If Len(clpBCode.Text) = 0 Then
    MsgBox "Benefit code is a required filed."
    clpBCode.SetFocus
    Exit Function
End If

If Len(clpBCode.Text) > 0 And clpBCode.Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpBCode.SetFocus
    Exit Function
End If

If Len(clpBGroup.Text) > 0 And clpBGroup.Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpBGroup.SetFocus
    Exit Function
End If


If Len(medLTSal(0)) < 1 Then
    MsgBox "You must have at least one Salary Range Entry."
    If medLTSal(0).Enabled Then medLTSal(0).SetFocus
    Exit Function
End If


Dim intRangesSet%
intRangesSet% = 0    ' 1 to 4 with 0 implying none
If Len(medLTSal(9)) = 0 Then
    medGTSal(9) = ""
Else
    If medLTSal(9) = 0 Then
        medLTSal(9) = ""
        medGTSal(9) = ""
    End If
End If


For X% = 0 To 9
    If Len(medLTSal(X%)) > 0 Then
        If Not IsNumeric(medLTSal(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medLTSal(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(medGTSal(X%)) > 0 Then
        If Not IsNumeric(medGTSal(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medGTSal(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(medPC(X%)) > 0 Then
        If Not IsNumeric(medPC(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medPC(X%).SetFocus
            Exit Function
        End If
    End If

    If Len(medLTSal(X%)) < 1 And Len(medGTSal(X%)) > 1 Then  ' missed one
        MsgBox "Ranges must be sequential"
        medLTSal(X%).SetFocus
        Exit Function
    End If
    If Len(medGTSal(X%)) > 0 Then
        If Val(medLTSal(X%)) > Val(medGTSal(X%)) Then
            MsgBox "Ranges must be sequential"
            medLTSal(X%).SetFocus
            Exit Function
        End If
    End If
    If X% > 0 And Len(medLTSal(X%)) > 0 Then
        If Val(medLTSal(X%)) < Val(medGTSal(X% - 1)) Then
            MsgBox "Ranges must be sequential"
            medLTSal(X%).SetFocus
            Exit Function
        End If
    End If
    If X% > 0 And Len(medGTSal(X%)) > 0 Then
        If Val(medGTSal(X%)) < Val(medGTSal(X% - 1)) And Val(medGTSal(X%)) <> 0 Then
            MsgBox "Ranges must be sequential"
            medLTSal(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(medLTSal(X%)) < 1 Then Exit For  ' missed one
    intRangesSet% = intRangesSet% + 1
Next X%
If intRangesSet% = 0 Then
    MsgBox "At least one Salary level must be set"
    medLTSal(0).SetFocus
    Exit Function
End If

chkMUBenCost = True

Exit Function

chkMUBenCost_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function



Sub cmdCancel_Click()
fglbNew = False
Data1.Refresh
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Call Display_Value
vbxTrueGrid.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub


Sub cmdDelete_Click()
Dim SQLQ, Msg, a%
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "The Benefit Costing details?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")
SQLQ = "DELETE FROM HR_BENEFIT_COST WHERE " & fglbBSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Data1.Refresh
Display_Value
End Sub

Sub cmdModify_Click()
OBGroup = clpBGroup.Text
OBCode = clpBCode.Text

fglbNew = False
End Sub

Sub cmdNew_Click()
Dim X
For X = 0 To 9
    medLTSal(X) = ""
    medGTSal(X) = ""
    medPC(X) = ""
    optM(X) = True
    optA(X) = False
    optW(X) = False
Next
clpBCode.Text = ""
clpBGroup.Text = ""
fglbNew = True
Call SET_UP_MODE
clpBGroup.SetFocus
End Sub

Sub cmdOK_Click()
Dim X%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsCU As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean

If Not chkMUBenCost() Then Exit Sub
For X% = 0 To 9
    If Not IsNumeric(medLTSal(X%)) Then Exit For
    If Not IsNumeric(medGTSal(X%)) Then
      medGTSal(X%) = 0
    Else
      If Val(medGTSal(X%)) = Int(medGTSal(X%)) Then medGTSal(X%) = medGTSal(X%) '+ 0.99
    End If
    If medLTSal(X%) > 0 And medGTSal(X%) = 0 Then medGTSal(X%) = 9999999
Next

If fglbNew Then
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HR_BENEFIT_COST WHERE " & fglbBSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
         clpBGroup.SetFocus
        Exit Sub
    End If
Else
    Call getWSQLQ("O")
    SQLQ = "DELETE FROM HR_BENEFIT_COST WHERE " & fglbBSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If
gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HR_BENEFIT_COST"
rsCU.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
For X% = 0 To 9
    If Len(medLTSal(X%)) > 0 Then
        rsCU.AddNew
        rsCU("CU_ORDER") = X + 1
        rsCU("CU_BENEFIT_GROUP_TABL") = "BGMF"
        rsCU("CU_BENEFIT_GROUP") = clpBGroup.Text
        rsCU("CU_BCODE_TABL") = "BNCD"
        rsCU("CU_BCODE") = clpBCode.Text
        rsCU("CU_MIN") = medLTSal(X%)
        rsCU("CU_MAX") = medGTSal(X%)
        If medPC(X%) = "" Then
            rsCU("CU_PCT") = Null
        Else
            rsCU("CU_PCT") = medPC(X%)
        End If
        If optM(X%) Then rsCU("CU_TYPE") = "M"
        If optA(X%) Then rsCU("CU_TYPE") = "A"
        'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
        If optW(X%) Then rsCU("CU_TYPE") = "W"
        rsCU("CU_FTE") = IIf(chkFTE(X%), 1, 0)
        rsCU.Update
    End If
Next
rsCU.Close
gdbAdoIhr001.CommitTrans
'If Not glbSQL and not glboracle Then Call Pause(0.5)
Call UPDBGroup(Trim(clpBGroup.Text), Trim(clpBCode.Text), "", GroupMasterEdit)
If glbCompSerial = "S/N - 2380W" Then Call CalcPP(Trim(clpBCode.Text), Trim(clpBGroup.Text))

Data1.Refresh
Display_Value

fglbNew = False
End Sub

Private Sub UPDBGroup(BGroup, BCode, Cover, BSource As BenefitUpdateSource)
Dim rsEMP As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim SQLQ
Dim xEMPNBR
Dim BCodeCover
Dim xTotalRecs, XUpdCount As Integer
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = "Update Benefit"
        
SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN (SELECT BF_EMPNBR FROM HRBENFT "
SQLQ = SQLQ & " WHERE BF_BCODE='" & BCode & "')"
If Len(BGroup) > 0 Then
    SQLQ = SQLQ & " AND ED_BENEFIT_GROUP='" & BGroup & "'"
End If
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
xTotalRecs = rsEMP.RecordCount

Do Until rsEMP.EOF
    MDIMain.panHelp(0).FloodPercent = (XUpdCount / xTotalRecs) * 100
    XUpdCount = XUpdCount + 1
    
    xEMPNBR = rsEMP("ED_EMPNBR")
    BCodeCover = BCode & "_" & Cover
    Call updateBenefit(xEMPNBR, BGroup, "A", BSource, BCodeCover)
    rsEMP.MoveNext
Loop
rsEMP.Close
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""

End Sub
Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Benefit Group Costing Details Report"
'Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 5
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgbcost.rpt"

SQLQ = "(1=1) "
If Len(clpBGroup.Text) > 0 Then SQLQ = SQLQ & " AND {HR_BENEFIT_COST.CU_BENEFIT_GROUP} = '" & clpBGroup.Text & "'"
If Len(clpBCode.Text) > 0 Then SQLQ = SQLQ & " AND {HR_BENEFIT_COST.CU_BCODE} = '" & clpBCode.Text & "'"
Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Benefit Group Costing Details Report"
'Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 5
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgbcost.rpt"

SQLQ = "(1=1) "
If Len(clpBGroup.Text) > 0 Then SQLQ = SQLQ & " AND {HR_BENEFIT_COST.CU_BENEFIT_GROUP} = '" & clpBGroup.Text & "'"
If Len(clpBCode.Text) > 0 Then SQLQ = SQLQ & " AND {HR_BENEFIT_COST.CU_BCODE} = '" & clpBCode.Text & "'"
Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Private Sub chkFTE_Click(Index As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Benefit Group Costing Details Report"
'Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 5
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgbcost.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub


Private Sub Form_Activate()
Call SET_UP_MODE
Call INI_Controls(Me)
glbOnTop = "FRMSBENCOST"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ
glbOnTop = "FRMSBENCOST"
Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT DISTINCT CU_BENEFIT_GROUP,CU_BCODE FROM HR_BENEFIT_COST "
Data1.RecordSource = SQLQ
Data1.Refresh
For X = 0 To 9
    If glbCompDecHR = 3 Then
        medLTSal(X).Format = "#,##0.000;(#,##0.000)"
        medGTSal(X).Format = "#,##0.000;(#,##0.000)"
    ElseIf glbCompDecHR = 4 Then
        medLTSal(X).Format = "#,##0.0000;(#,##0.0000)"
        medGTSal(X).Format = "#,##0.0000;(#,##0.0000)"
    Else
        medLTSal(X).Format = "#,##0.00;(#,##0.00)"
        medGTSal(X).Format = "#,##0.00;(#,##0.00)"
    End If
Next
Call INI_Controls(Me)
ST_UPD_MODE (False)
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim Keepfocus As Boolean
'If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
'Keepfocus = Not isUpdated(Me)
'Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
If Me.Height >= 2880 + VacFram.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    VacFram.Top = 2880
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = VacFram.Height + panControls.Height + 2880 + 550 - Me.Height
scrControl.Left = Me.Width - scrControl.Width - 120
If Me.Height - scrControl.Top - panControls.Height - 300 > 0 Then
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 300
Else
    scrControl.Height = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."

Set frmSBenCost = Nothing  'carmen apr 2000
End Sub



Private Sub medPC_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
If Len(medPC(Index)) > 0 Then
    medPC(Index) = medPC(Index) * 100
End If
End Sub

Private Sub medPC_LostFocus(Index As Integer)
If (Not IsNumeric(medPC(Index))) And medPC(Index).DataChanged Then medPC(Index) = 0
If Len(medPC(Index)) > 0 Then
    medPC(Index) = medPC(Index) / 100
End If
End Sub

Private Sub medGTSal_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medlTSal_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub






Private Sub optM_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optM_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optA_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optA_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optW_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optW_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
VacFram.Top = 3750 - scrControl.Value
End Sub

Sub ST_UPD_MODE(TF As Boolean)
Dim X, FT
FT = Not TF
For X = 0 To 9
    medLTSal(X).Enabled = TF
    medGTSal(X).Enabled = TF
    medPC(X).Enabled = TF
    If X = 0 Then
        optM(X).Enabled = TF
        optA(X).Enabled = TF
        optW(X).Enabled = TF
        chkFTE(X).Enabled = TF
    Else
        optM(X).Enabled = False
        optA(X).Enabled = False
        optW(X).Enabled = False
        chkFTE(X).Enabled = False
    End If
Next

clpBGroup.Enabled = TF
clpBCode.Enabled = TF

End Sub

Private Sub getWSQLQ(xType)
Dim xBGroup, xBCode

If xType = "" Then Exit Sub

If xType = "O" Then
    xBGroup = OBGroup
    xBCode = OBCode
Else
    xBGroup = clpBGroup.Text
    xBCode = clpBCode.Text
End If

fglbBSQLQ = " CU_BCODE ='" & xBCode & "'"
If Len(xBGroup) = 0 Then
    fglbBSQLQ = fglbBSQLQ & " AND (CU_BENEFIT_GROUP IS NULL OR CU_BENEFIT_GROUP='')"
Else
    fglbBSQLQ = fglbBSQLQ & " AND CU_BENEFIT_GROUP = '" & xBGroup & "'"
End If


End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsCU As New ADODB.Recordset
Dim X
For X = 0 To 9
    medLTSal(X) = ""
    medGTSal(X) = ""
    medPC(X) = ""
    optM(X) = True
    optA(X) = False
    optW(X) = False
    chkFTE(X).Value = 0
Next
clpBGroup.Text = ""
clpBCode.Text = ""


If Not Data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HR_BENEFIT_COST "
    If IsNull(Data1.Recordset("CU_BENEFIT_GROUP")) Then
        SQLQ = SQLQ & " WHERE CU_BENEFIT_GROUP IS NULL"
    Else
        SQLQ = SQLQ & " WHERE CU_BENEFIT_GROUP = '" & Data1.Recordset("CU_BENEFIT_GROUP") & "'"
    End If
    SQLQ = SQLQ & " AND CU_BCODE = '" & Data1.Recordset("CU_BCODE") & "'"
    
    
    
    SQLQ = SQLQ & " Order By CU_BENEFIT_GROUP ,CU_BCODE,CU_ORDER "
    rsCU.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not IsNull(Data1.Recordset("CU_BENEFIT_GROUP")) Then clpBGroup.Text = Data1.Recordset("CU_BENEFIT_GROUP")
    If Not IsNull(Data1.Recordset("CU_BCODE")) Then clpBCode.Text = Data1.Recordset("CU_BCODE")
    
    
    Do While Not rsCU.EOF
        xOrder = rsCU("CU_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 16) Then
            If Not IsNull(rsCU("CU_MIN")) Then medLTSal(nOrder) = rsCU("CU_MIN")
            If Not IsNull(rsCU("CU_MAX")) Then medGTSal(nOrder) = rsCU("CU_MAX")
            If Not IsNull(rsCU("CU_PCT")) Then medPC(nOrder) = rsCU("CU_PCT")
            If rsCU("CU_TYPE") = "M" Then optM(nOrder) = True
            If rsCU("CU_TYPE") = "A" Then optA(nOrder) = True
            'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
            If rsCU("CU_TYPE") = "W" Then optW(nOrder) = True
            If rsCU("CU_FTE") Then chkFTE(nOrder) = 1
        End If
        rsCU.MoveNext
    Loop
    rsCU.Close
End If
SET_UP_MODE
Call cmdModify_Click
End Sub



Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT DISTINCT CU_BENEFIT_GROUP,CU_BCODE FROM HR_BENEFIT_COST "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub


Sub ST_OPT_VALUE()
Dim X, XoptM, XoptA, XoptF, XoptW
    XoptM = optM(0).Value
    XoptA = optA(0).Value
    XoptW = optW(0).Value
    XoptF = chkFTE(0).Value
    For X = 1 To 9
        optM(X).Value = XoptM
        optA(X).Value = XoptA
        optW(X).Value = XoptW
        chkFTE(X).Value = XoptF
    Next
End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
    cmdPrintAll.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)

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
RelateMode = NothingRelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_BenefitGroupSetup 'gSec_Upd_Entitlements
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

Private Sub CalcPP(Optional xCode As String, Optional xGroup As String)
    Dim rs As New ADODB.Recordset
    Dim rsIn As New ADODB.Recordset
    Dim SQLQ As String, WSQLQ As String
    Dim X As Boolean
    Dim I As Long, xTot As Long, oPayP As Double
    
    WSQLQ = ""
    If IsEmpty(xCode) = False Then
        If xCode <> "" Then
            WSQLQ = "WHERE BF_BCODE='" & xCode & "' and BF_GROUP='" & xGroup & "'"
        End If
    End If
    
    SQLQ = "SELECT BF_EMPNBR, BF_PPAMT, BF_MTHECOST, BF_GROUP, BF_BCODE, BF_LUSER, BF_LDATE, BF_LTIME FROM HRBENFT "
    SQLQ = SQLQ & WSQLQ
    
    rs.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not rs.EOF Then
        xTot = rs.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    Do While Not rs.EOF
    'If rs.EOF = False And rs.BOF = False Then
        MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100: I = I + 1
        oPayP = rs("BF_PPAMT")
        rs("BF_PPAMT") = rs("BF_MTHECOST") / 2
        rs("BF_LUSER") = glbUserID
        rs("BF_LTIME") = Time$
        rs("BF_LDATE") = Format(Now, "SHORT DATE")
        rs.Update
        If oPayP <> rs("BF_PPAMT") Then
            X = AUDITPP(rs("BF_EMPNBR"), rs("BF_BCODE"), rs("BF_MTHECOST") / 2)
        End If
        rs.MoveNext
    'End If
    Loop
    rs.Close

    MDIMain.panHelp(0).FloodType = 0
End Sub




Private Function AUDITPP(xEMPNBR, xCode, xAmount) As Boolean
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, ACTX As String
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITPP = False
ACTX = "M"
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEMPNBR, gdbAdoIhr001, adOpenKeyset

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

'rsTB.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenKeyset, adCmdText
'If rsTB.EOF Then GoTo MODNOUPD

'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_PPAMT") = xAmount 'rsTB("BF_PPAMT") AU_BCODE
rsTA("AU_BCODE") = xCode

Dim rsEMP As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEMP.EOF Then
    If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
End If
rsEMP.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEMPNBR 'glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITPP = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function


