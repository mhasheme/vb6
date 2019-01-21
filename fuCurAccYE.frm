VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUCurAccYEnd 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Current Accrued Year End Update"
   ClientHeight    =   7290
   ClientLeft      =   2565
   ClientTop       =   -1140
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
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7290
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdYearEnd 
      Appearance      =   0  'Flat
      Caption         =   "Current Accrued Year End Update"
      Height          =   375
      Left            =   3773
      TabIndex        =   2
      Top             =   2880
      Width           =   4065
   End
   Begin VB.Frame VacFram02 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   840
      TabIndex        =   62
      Top             =   8520
      Visible         =   0   'False
      Width           =   7815
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   10
         Left            =   9390
         TabIndex        =   63
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   11
         Left            =   9390
         TabIndex        =   64
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   405
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   12
         Left            =   9390
         TabIndex        =   65
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   720
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   13
         Left            =   9390
         TabIndex        =   66
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1050
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   14
         Left            =   9390
         TabIndex        =   67
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1365
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   68
         Tag             =   "11-Service is greater than this number"
         Top             =   60
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   10
         Left            =   1965
         TabIndex        =   69
         Tag             =   "10-Service is less than this number"
         Top             =   60
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   70
         Tag             =   "11-Service is greater than this number"
         Top             =   390
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   11
         Left            =   1965
         TabIndex        =   71
         Tag             =   "10-Service is less than this number"
         Top             =   390
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   10
         Left            =   3690
         TabIndex        =   72
         Tag             =   "11-Entitlement Amount"
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   11
         Left            =   3690
         TabIndex        =   73
         Tag             =   "11-Entitlement Amount"
         Top             =   405
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   10
         Left            =   4890
         TabIndex        =   74
         Top             =   0
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   10
            Left            =   1770
            TabIndex        =   75
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   76
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   10
            Left            =   930
            TabIndex        =   77
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   11
         Left            =   4890
         TabIndex        =   78
         Top             =   330
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   11
            Left            =   1770
            TabIndex        =   79
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   80
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   11
            Left            =   930
            TabIndex        =   81
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   10
         Left            =   7815
         TabIndex        =   82
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   11
         Left            =   7815
         TabIndex        =   83
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   405
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   84
         Tag             =   "11-Service is greater than this number"
         Top             =   720
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   12
         Left            =   1965
         TabIndex        =   85
         Tag             =   "10-Service is less than this number"
         Top             =   720
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   86
         Tag             =   "11-Service is greater than this number"
         Top             =   1035
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   13
         Left            =   1965
         TabIndex        =   87
         Tag             =   "10-Service is less than this number"
         Top             =   1035
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   14
         Left            =   0
         TabIndex        =   88
         Tag             =   "11-Service is greater than this number"
         Top             =   1350
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   14
         Left            =   1965
         TabIndex        =   89
         Tag             =   "10-Service is less than this number"
         Top             =   1350
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   12
         Left            =   3690
         TabIndex        =   90
         Tag             =   "11-Entitlement Amount"
         Top             =   720
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   13
         Left            =   3690
         TabIndex        =   91
         Tag             =   "11-Entitlement Amount"
         Top             =   1035
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   12
         Left            =   4890
         TabIndex        =   92
         Top             =   660
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   12
            Left            =   1770
            TabIndex        =   93
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   12
            Left            =   930
            TabIndex        =   94
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   12
            Left            =   90
            TabIndex        =   95
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   13
         Left            =   4890
         TabIndex        =   96
         Top             =   975
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   13
            Left            =   1770
            TabIndex        =   97
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   98
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   13
            Left            =   930
            TabIndex        =   99
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   14
         Left            =   4890
         TabIndex        =   100
         Top             =   1275
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   14
            Left            =   1770
            TabIndex        =   101
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   102
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   14
            Left            =   930
            TabIndex        =   103
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   12
         Left            =   7815
         TabIndex        =   104
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   720
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   13
         Left            =   7815
         TabIndex        =   105
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1050
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   14
         Left            =   7815
         TabIndex        =   106
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1365
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   14
         Left            =   3690
         TabIndex        =   107
         Tag             =   "11-Entitlement Amount"
         Top             =   1350
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   15
         Left            =   9390
         TabIndex        =   108
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1700
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   16
         Left            =   9390
         TabIndex        =   109
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2020
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   110
         Tag             =   "11-Service is greater than this number"
         Top             =   1700
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   15
         Left            =   1965
         TabIndex        =   111
         Tag             =   "10-Service is less than this number"
         Top             =   1700
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   112
         Tag             =   "11-Service is greater than this number"
         Top             =   2020
         Width           =   540
         _ExtentX        =   953
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   16
         Left            =   1965
         TabIndex        =   113
         Tag             =   "10-Service is less than this number"
         Top             =   2020
         Width           =   540
         _ExtentX        =   953
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   15
         Left            =   3690
         TabIndex        =   114
         Tag             =   "11-Entitlement Amount"
         Top             =   1700
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   15
         Left            =   4890
         TabIndex        =   115
         Top             =   1600
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   15
            Left            =   1770
            TabIndex        =   116
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   117
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   15
            Left            =   930
            TabIndex        =   118
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   16
         Left            =   4890
         TabIndex        =   119
         Top             =   1940
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   16
            Left            =   1770
            TabIndex        =   120
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
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
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   121
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   16
            Left            =   930
            TabIndex        =   122
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
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
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   15
         Left            =   7815
         TabIndex        =   123
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1700
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   16
         Left            =   7815
         TabIndex        =   124
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   2020
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   16
         Left            =   3690
         TabIndex        =   125
         Tag             =   "11-Entitlement Amount"
         Top             =   2020
         Width           =   870
         _ExtentX        =   1535
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   510
         TabIndex        =   132
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   510
         TabIndex        =   131
         Top             =   1065
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   510
         TabIndex        =   130
         Top             =   1365
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   510
         TabIndex        =   129
         Top             =   105
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   510
         TabIndex        =   128
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   510
         TabIndex        =   127
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   480
         TabIndex        =   126
         Top             =   2040
         Width           =   1440
      End
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   4
      Left            =   2820
      TabIndex        =   1
      Tag             =   "11-Service is greater than this number"
      Top             =   9285
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   4
      Left            =   4785
      TabIndex        =   3
      Tag             =   "10-Service is less than this number"
      Top             =   9285
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   5
      Left            =   2820
      TabIndex        =   4
      Tag             =   "11-Service is greater than this number"
      Top             =   9615
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   5
      Left            =   4785
      TabIndex        =   5
      Tag             =   "10-Service is less than this number"
      Top             =   9615
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   6
      Left            =   2820
      TabIndex        =   6
      Tag             =   "11-Service is greater than this number"
      Top             =   9960
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   6
      Left            =   4785
      TabIndex        =   7
      Tag             =   "10-Service is less than this number"
      Top             =   9960
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   7
      Left            =   2820
      TabIndex        =   8
      Tag             =   "11-Service is greater than this number"
      Top             =   10320
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   7
      Left            =   4785
      TabIndex        =   9
      Tag             =   "10-Service is less than this number"
      Top             =   10335
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   4
      Left            =   12210
      TabIndex        =   10
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   9270
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   4
      Left            =   6510
      TabIndex        =   11
      Tag             =   "11-Entitlement Amount"
      Top             =   9630
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   5
      Left            =   6510
      TabIndex        =   12
      Tag             =   "11-Entitlement Amount"
      Top             =   9975
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   6
      Left            =   6510
      TabIndex        =   13
      Tag             =   "11-Entitlement Amount"
      Top             =   10335
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   4
      Left            =   7710
      TabIndex        =   14
      Top             =   9180
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   4
         Left            =   1785
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
      Begin Threed.SSOption optD 
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Days"
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
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   5
      Left            =   7710
      TabIndex        =   18
      Top             =   9540
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   5
         Left            =   1770
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optD 
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   20
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Days"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   21
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   6
      Left            =   7710
      TabIndex        =   22
      Top             =   9870
      Visible         =   0   'False
      Width           =   2640
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   6
         Left            =   1770
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optD 
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Days"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   25
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   7
      Left            =   7710
      TabIndex        =   26
      Top             =   10200
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   7
         Left            =   1785
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optD 
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   28
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Days"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   7
         Left            =   945
         TabIndex        =   29
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   4
      Left            =   10635
      TabIndex        =   30
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   9285
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   5
      Left            =   10635
      TabIndex        =   31
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   9630
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   6
      Left            =   10635
      TabIndex        =   32
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   9975
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   7
      Left            =   10635
      TabIndex        =   33
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   10335
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   7
      Left            =   6510
      TabIndex        =   34
      Tag             =   "11-Entitlement Amount"
      Top             =   9270
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   5
      Left            =   12210
      TabIndex        =   35
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   9630
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   6
      Left            =   12210
      TabIndex        =   36
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   9975
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   7
      Left            =   12210
      TabIndex        =   37
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   10335
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   8
      Left            =   2820
      TabIndex        =   42
      Tag             =   "11-Service is greater than this number"
      Top             =   10695
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   8
      Left            =   4785
      TabIndex        =   43
      Tag             =   "10-Service is less than this number"
      Top             =   10695
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   9
      Left            =   2820
      TabIndex        =   44
      Tag             =   "11-Service is greater than this number"
      Top             =   11025
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medGTServ 
      Height          =   285
      Index           =   9
      Left            =   4785
      TabIndex        =   45
      Tag             =   "10-Service is less than this number"
      Top             =   11025
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   8
      Left            =   12210
      TabIndex        =   46
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   10680
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   8
      Left            =   6510
      TabIndex        =   47
      Tag             =   "11-Entitlement Amount"
      Top             =   11040
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   8
      Left            =   7710
      TabIndex        =   48
      Top             =   10590
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   8
         Left            =   1785
         TabIndex        =   49
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   50
         TabStop         =   0   'False
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
      Begin Threed.SSOption optD 
         Height          =   225
         Index           =   8
         Left            =   90
         TabIndex        =   51
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Days"
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
   End
   Begin Threed.SSFrame frmDH 
      Height          =   375
      Index           =   9
      Left            =   7710
      TabIndex        =   52
      Top             =   10950
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optF 
         Height          =   195
         Index           =   9
         Left            =   1770
         TabIndex        =   53
         TabStop         =   0   'False
         Tag             =   "Entitlement Measured in FTE#"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "FTE#"
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
      Begin Threed.SSOption optD 
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   54
         Tag             =   "Entitlement measured in days"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Days"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optH 
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   55
         Tag             =   "Entitlement measured in hours"
         Top             =   120
         Width           =   750
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Hours"
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
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   8
      Left            =   10635
      TabIndex        =   56
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   10695
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMax 
      Height          =   285
      Index           =   9
      Left            =   10635
      TabIndex        =   57
      Tag             =   "10-Maximum Entitlement can be"
      Top             =   11040
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEntitle 
      Height          =   285
      Index           =   9
      Left            =   6510
      TabIndex        =   58
      Tag             =   "11-Entitlement Amount"
      Top             =   10680
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medVacation 
      Height          =   285
      Index           =   9
      Left            =   12210
      TabIndex        =   59
      Tag             =   "10-Vacation Pay Percentage"
      Top             =   11040
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   675
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      Left            =   600
      TabIndex        =   135
      Top             =   720
      Width           =   555
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
      Left            =   480
      TabIndex        =   134
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Make sure all the Vacation Time Taken entries have been entered in Attendance for the Current Year."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   480
      TabIndex        =   133
      Top             =   2160
      Width           =   10440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<=  Service  =>"
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
      Left            =   3330
      TabIndex        =   61
      Top             =   10710
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<=  Service  =>"
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
      Left            =   3330
      TabIndex        =   60
      Top             =   11055
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<=  Service  =>"
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
      Left            =   3330
      TabIndex        =   41
      Top             =   9300
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<=  Service  =>"
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
      Left            =   3330
      TabIndex        =   40
      Top             =   9645
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<=  Service  =>"
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
      Left            =   3330
      TabIndex        =   39
      Top             =   9975
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblSer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">    Service  "
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
      Left            =   3525
      TabIndex        =   38
      Top             =   10320
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmUCurAccYEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim snapOvertime As New ADODB.Recordset
Dim fglbWDate$, fglbWDateS$
Dim fglbModifyWhat$
Dim fglbSick%
Dim fglbVac%

Dim fglbSDate As Variant
Dim fglbMaxRange%
Dim fglbCompMonthly%

Dim ffieldEntitle$    ' ED_VAC or ED_SICK for name of field for entitlement
Dim ffieldPEntitle$     ' ED_PVAC or ED_PSICK for previous entitlement's field name
Dim ffieldTEntitle$     ' ED_VACT or ED_SICKT for taken entitlement's field name For whscc Sick Update
Dim fglbCode$           ' are we dealing with Vac/Sick records?"
Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&
Dim UpdVac As Boolean
Dim ControlsShown As Boolean
Dim fglbESQLQ
Dim Memplist1, Memplist2

Public Sub cmdClose_Click()
Unload Me

End Sub

Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%

On Error GoTo Mod_Err
If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
'Screen.MousePointer = HOURGLASS
'
'Select Case fglbModifyWhat$
'    Case "Zero Out"
'        If Not modZeroSelection() Then Exit Sub
'    Case "Rollover"
'        If Not modRollSelection() Then Exit Sub
'End Select
'
'Call EntReCalc(fglbESQLQ)
'Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub


'Public Sub cmdRollover_Click()
''panJGroup.Visible = False       'js-6Apr99
'
'ControlsShown = False
'Call UpdateEntControls(ControlsShown)
'fglbModifyWhat$ = "Rollover"
'
'frmZero.Visible = False '
'frmRoll.Visible = True
'frmRoll.Top = 4000
'lblRollEntitlements.Visible = True
'lblRollEntitlements.Left = 120
'lblRollEntitlements.Top = 3300
'Me.Caption = "Rollover Entitlements"
'
''Town of Aurora
''If glbCompSerial = "S/N - 2378W" Then
'    chkOvtE.Visible = True
'    chkSick.SetFocus
'    chkSick.Value = True
''Else
''    optOvtE.Visible = False
''End If
'
'End Sub

'Public Sub cmdZeroOut_Click()
''optBothVS.Visible = True
'
'fglbModifyWhat$ = "Zero Out"
'ControlsShown = False
'Call UpdateEntControls(ControlsShown)
'
'lblRollEntitlements.Visible = False
'lblZero.Visible = True 'js-6Apr99
'lblZero.Left = 120     '
'lblZero.Top = 3300     '
'
'frmZero.Visible = True '
'frmZero.Top = 4000
'
'frmRoll.Visible = False
'Me.Caption = "Zero Out Entitlements"
'
''Town of Aurora
''If glbCompSerial = "S/N - 2378W" Then
'    chkOvtE.Visible = True
'    chkSick.SetFocus
'    chkSick.Value = True
''Else
''    optOvtE.Visible = False
''End If
'
'End Sub

Private Function CR_SnapEntitle()
Dim SQLQ As String, SQLQ1 As String
Dim snapMultiEmp As New ADODB.Recordset

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ
'Multi Positions Update #3304
'gdbAdoIhr001.CursorLocation = adUseServer
''' Please talk to Jaddy
'''If glbMulti Then
'''    Memplist1 = "": Memplist2 = ""
'''    If UCase(glbCompEntVac$) = "A" Or UCase(glbCompEntSick$) = "A" Then
'''    If glbOracle Then
'''        SQLQ1 = "SELECT HREMP.ED_EMPNBR, COUNT(ED_EMPNBR) AS SUMEMP "
'''        SQLQ1 = SQLQ1 & " FROM HREMP , qry_JobCurrent WHERE HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
'''        SQLQ1 = SQLQ1 & "AND " & fglbESQLQ
'''    Else
'''        SQLQ1 = "SELECT HREMP.ED_EMPNBR, COUNT(ED_EMPNBR) AS SUMEMP "
'''        SQLQ1 = SQLQ1 & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
'''        SQLQ1 = SQLQ1 & "WHERE " & fglbESQLQ
'''    End If
'''        If Len(clpCode(4).Text) > 0 Then
'''            SQLQ1 = SQLQ1 & " AND ED_EMPNBR IN "
'''            SQLQ1 = SQLQ1 & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'''            SQLQ1 = SQLQ1 & " WHERE JB_GRPCD = '" & clpCode(4).Text & "') "
'''        End If
'''        If snapMultiEmp.State <> 0 Then snapMultiEmp.Close
'''        SQLQ1 = SQLQ1 & " GROUP BY ED_EMPNBR HAVING COUNT(ED_EMPNBR) > 1 "
'''        snapMultiEmp.Open SQLQ1, gdbAdoIhr001, adOpenStatic
'''        Do While Not snapMultiEmp.EOF
'''            Memplist1 = Memplist1 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
'''            Memplist2 = Memplist2 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
'''            snapMultiEmp.MoveNext
'''        Loop
'''        snapMultiEmp.Close
'''    End If
'''End If
'Multi Positions Update #3304

'Jaddy Changed. Cound not join the position table it will cost so many problems.

SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_VACT, ED_ANNSICK, ED_ANNVAC, "
SQLQ = SQLQ & " ED_DIV,ED_PT, "
SQLQ = SQLQ & " ED_DEPTNO,ED_ORG, " 'NEW BY Frank for County of Elgin ticket #4653
SQLQ = SQLQ & " ED_LOC,ED_SECTION,ED_SALDIST, "  'NEW
SQLQ = SQLQ & " ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES, "  'NEW
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME"
SQLQ = SQLQ & " FROM HREMP "
SQLQ = SQLQ & " WHERE " & fglbESQLQ
'If Len(clpCode(4).Text) > 0 Then
'    SQLQ = SQLQ & " AND ED_EMPNBR IN "
'    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(4).Text & "') "
'End If
If snapEntitle.State <> 0 Then snapEntitle.Close
If glbOracle Then
    snapEntitle.CursorLocation = adUseServer
End If
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic  'adLockPessimistic ''

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

Private Sub cmdYearEnd_Click()
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim Response%, pct%, prec%, xErr
    Dim X%, DtTm As Variant, Msg$, Title$, DgDef As Variant
    Dim xCurrVac
    Dim lngRecs&
    
    On Error GoTo modCurrYearEnd_Err
    
    Screen.MousePointer = HOURGLASS
    
    Call getWSQLQ   'Department Security & selection criteria
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    
    SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_VACT, ED_ETDATE, ED_LDATE, ED_LTIME, ED_LUSER FROM HREMP WHERE " & fglbESQLQ
    'SQLQ = SQLQ & " AND ED_EMPNBR = 41332"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsHREmp.BOF And rsHREmp.EOF Then
        MsgBox "No employees to update! Check the Department Security setting."
        Exit Sub
    Else
        lngRecs& = rsHREmp.RecordCount
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
        Title$ = "Current Accrued Year End Update"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            MDIMain.panHelp(0).FloodType = 0
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    
        gdbAdoIhr001.BeginTrans
    
        rsHREmp.MoveFirst
        Do While Not rsHREmp.EOF
            prec% = prec% + 1
            pct% = Int(100 * (prec% / (lngRecs&)))
            MDIMain.panHelp(0).FloodPercent = pct%
        
            xCurrVac = IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC"))
            
            rsHREmp("ED_VAC") = IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC")) - IIf(IsNull(rsHREmp("ED_VACT")), 0, rsHREmp("ED_VACT"))
            rsHREmp("ED_VACT") = 0
            rsHREmp("ED_LDATE") = Now
            rsHREmp("ED_LTIME") = Time$
            rsHREmp("ED_LUSER") = glbLEE_ID
            
        
            'xComments = "Current Vac. Accrued Chg from " & xCurrVac & " to " & rsHREmp("ED_VAC")
            'Call Append_Accrual(EmpNo&, "VAC", rsHREmp("ED_ETDATE"), Val(rsHREmp("ED_VAC") - xCurrVac & ""), "U", xComments)
            
            rsHREmp.Update
            
            rsHREmp.MoveNext
        Loop
    End If
    gdbAdoIhr001.CommitTrans
    
    rsHREmp.Close
    
    If lngRecs& > 0 Then
        MsgBox "Current Accrued Year End Update completed successfully.", vbOKOnly, "Current Accrued Year End Update"
    End If
    
    MDIMain.panHelp(0).FloodType = 0
    Screen.MousePointer = DEFAULT
    Unload Me
Exit Sub

modCurrYearEnd_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Current Accrued Year End", "HREMP", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUCURACCYEND"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUCURACCYEND"

'MDIMain.lstPanel.Visible = False
'MDIMain.lstView.Visible = False

'Dim rsVE As New ADODB.Recordset
'chkSick = True  'default to sick
'Select Case glbCompWDate$ ' sets field reference for basic 'which date'
'    Case "O": fglbWDate$ = "ED_DOH"
'    Case "S": fglbWDate$ = "ED_SENDTE"
'    Case "U": fglbWDate$ = "ED_UNION"
'    Case "L": fglbWDate$ = "ED_LTHIRE"
'    Case "D": fglbWDate$ = "ED_USRDAT1"
'End Select
'
'Screen.MousePointer = HOURGLASS
'Call setRptCaption(Me)
'UpdVac = False
'
'Call modSetFGlobals("Sick")
'
'If glbMulti Then textMulti.Visible = True
'If glbLinamar Then
'    lblSection = "Vacation Group"
'    clpCode(1).LookupType = SalaryDistribution
'
'End If
'
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
'If glbUEnt = 2 Then
'    Me.Caption = "Rollover Entitlements"
'Else
'    Me.Caption = "Zero Out Entitlements"
'End If

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."
glbUEnt = 0
Set frmUCurAccYEnd = Nothing  'carmen apr 2000
End Sub

Private Function modRollSelection()
Dim EmpNo As Long, strJob$, spt As Variant, lngRecs&
Dim X%, DtTm As Variant, Msg$, Title$, DgDef As Variant
Dim Response%, pct%, prec%, xErr
Dim SQLQ As String, dblOUTS#, dblOUTV#
On Error GoTo modRollSelection_Err
modRollSelection = False

xErr = False
Msg$ = ""
'If (chkSick) And glbEntOutStandingS$ <> "1" Then
'    Msg$ = Msg$ & "SickTime Oustanding Entitlements" & Chr(10) & "is not based on Entitlement Date" & Chr(10)
'    xErr = True
'End If
'If (chkVacation) And glbEntOutStanding$ <> "1" Then
'    Msg$ = Msg$ & "Vacation Oustanding Entitlements" & Chr(10) & "is not based on Entitlement Date" & Chr(10)
'    xErr = True
'End If
If xErr Then
    Title$ = "RollOver ERROR"
    Msg$ = Msg$ & "ERROR - Cannot Continue !!!"
    DgDef = MB_ICONEXCLAMATION
    MsgBox Msg$, DgDef, Title$
    Exit Function
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
'    If chkOvtE Then
'        Call Overtime_Bank_Rollover
'    End If
'End If
'If chkVacation Or chkSick Then

    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
    
    Screen.MousePointer = DEFAULT
    
    If snapEntitle.BOF And snapEntitle.EOF Then
        MsgBox "Employees for this selection do not exist!"
        Exit Function
    Else
        lngRecs& = snapEntitle.RecordCount
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
        Title$ = "Update Entitlements"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Function
        End If
    End If
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    Screen.MousePointer = HOURGLASS
    
    Dim xComments
    
    gdbAdoIhr001.BeginTrans
    
    While Not snapEntitle.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        EmpNo& = snapEntitle("ED_EMPNBR")
        
        dblOUTS# = 0
        'If chkSick Then
            If IsNumeric(snapEntitle("ED_PSICK")) Then
                  dblOUTS# = dblOUTS# + snapEntitle("ED_PSICK")
            End If
            If IsNumeric(snapEntitle("ED_SICK")) Then
                  dblOUTS# = dblOUTS# + snapEntitle("ED_SICK")
            End If
            If IsNumeric(snapEntitle("ED_SICKT")) Then
                  dblOUTS# = dblOUTS# - snapEntitle("ED_SICKT")
            End If
        'End If
        'Frank 10/21/03 ticket #2292
        'If Dept. = 2 or 11 and Union = 2, Maximum previous year (rollover) can be 16 hours.
        If glbCElgin Then
            If (snapEntitle("ED_DEPTNO") = "2" Or snapEntitle("ED_DEPTNO") = "11") And snapEntitle("ED_ORG") = "2" Then
                If dblOUTS# > 16 Then dblOUTS# = 16
            End If
        End If
        dblOUTV# = 0
        'If chkVacation Then
            If IsNumeric(snapEntitle("ED_PVAC")) Then
                  dblOUTV# = dblOUTV# + snapEntitle("ED_PVAC")
            End If
            If IsNumeric(snapEntitle("ED_VAC")) Then
                  dblOUTV# = dblOUTV# + snapEntitle("ED_VAC")
            End If
            If IsNumeric(snapEntitle("ED_VACT")) Then
                  dblOUTV# = dblOUTV# - snapEntitle("ED_VACT")
            End If
        'End If
    
        'If chkSick = True Then
            'If optAccumulate Then
                If dblOUTS# < 0 Then
                    xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & dblOUTS#
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "SICK", Date, dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    Call Append_Accrual(EmpNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PSICK") = dblOUTS#
                Else
                    xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & 0
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "SICK", Date, 0 - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    Call Append_Accrual(EmpNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PSICK") = 0
                End If
            'Else
                xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & dblOUTS#
                '================= By Hemu
                'Call Append_Accrual(EmpNo&, "SICK", Date, dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                Call Append_Accrual(EmpNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                '================= By Hemu
                snapEntitle("ED_PSICK") = dblOUTS#
            'End If
        'End If
        
        'If chkVacation Then
            'If optAccumulate Then
                If dblOUTV# < 0 Then
                    xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & dblOUTV#
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "VAC", Date, dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    Call Append_Accrual(EmpNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PVAC") = dblOUTV#
                Else
                    xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & 0
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "VAC", Date, 0 - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    Call Append_Accrual(EmpNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PVAC") = 0
                End If
            'Else
                xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & dblOUTV#
                '================= By Hemu
                'Call Append_Accrual(EmpNo&, "VAC", Date, dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                Call Append_Accrual(EmpNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                '================= By Hemu
                snapEntitle("ED_PVAC") = dblOUTV#
            'End If
        'End If
        snapEntitle("ED_LDATE") = Now
        snapEntitle("ED_LTIME") = Time$
        snapEntitle("ED_LUSER") = glbLEE_ID
        snapEntitle.Update
        snapEntitle.MoveNext
    Wend
    gdbAdoIhr001.CommitTrans
    modRollSelection = True
    MDIMain.panHelp(0).FloodType = 0
    snapEntitle.Close
    Screen.MousePointer = DEFAULT
'End If

Exit Function

modRollSelection_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Roll-Over Entitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Overtime_Bank_Rollover()
Dim EmpNo As Long, strJob$, spt As Variant, lngRecs&
Dim X%, DtTm As Variant, Msg$, Title$, DgDef As Variant
Dim Response%, pct%, prec%, xErr
Dim SQLQ As String, dblOUTO
Dim xComments

On Error GoTo Overtime_Bank_Rollover_Err


If Not CR_SnapOvertime() Then Exit Sub  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT

If snapOvertime.BOF And snapOvertime.EOF Then
    MsgBox "Employees for Overtime Bank do not exist!"
    Exit Sub
Else
    lngRecs& = snapOvertime.RecordCount
    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
    Title$ = "Rollover Overtime Bank"
    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5
Screen.MousePointer = HOURGLASS

gdbAdoIhr001.BeginTrans

While Not snapOvertime.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    EmpNo& = snapOvertime("OT_EMPNBR")
    
    dblOUTO = 0
    'If chkOvtE Then
        If IsNumeric(snapOvertime("OT_PBANK")) Then
              dblOUTO = dblOUTO + snapOvertime("OT_PBANK")
        End If
        If IsNumeric(snapOvertime("OT_BANK")) Then
              dblOUTO = dblOUTO + snapOvertime("OT_BANK")
        End If
        If IsNumeric(snapOvertime("OT_BANKT")) Then
              dblOUTO = dblOUTO - snapOvertime("OT_BANKT")
        End If
    
        'If optAccumulate Then
            If dblOUTO < 0 Then
                xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & dblOUTO
                Call Append_Accrual(EmpNo&, "BANK", Date, dblOUTO - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
                snapOvertime("OT_PBANK") = dblOUTO
            Else
                xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & 0
                Call Append_Accrual(EmpNo&, "BANK", Date, 0 - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
                snapOvertime("OT_PBANK") = 0
            End If
        'Else
            xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & dblOUTO
            Call Append_Accrual(EmpNo&, "BANK", Date, dblOUTO - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
            snapOvertime("OT_PBANK") = dblOUTO
        'End If
    'End If
    snapOvertime("OT_LDATE") = Now
    snapOvertime("OT_LTIME") = Time$
    snapOvertime("OT_LUSER") = glbLEE_ID
    snapOvertime.Update
    snapOvertime.MoveNext
Wend
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodType = 0
snapOvertime.Close
Screen.MousePointer = DEFAULT

Exit Sub

Overtime_Bank_Rollover_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Roll-Over Bank", "HR_OVERTIME_BANK", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Overtime_Bank_ZeroOut()
Dim EmpNo As Long, strJob$, spt As Variant, lngRecs&
Dim X%, DtTm As Variant, Msg$, Title$, DgDef As Variant
Dim Response%, pct%, prec%, xErr
Dim SQLQ As String, dblOUTO
Dim xComments

On Error GoTo Overtime_Bank_ZeroOut_Err

If Not CR_SnapOvertime() Then Exit Sub  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT

If snapOvertime.BOF And snapOvertime.EOF Then
    MsgBox "Employees for Overtime Bank do not exist!"
    Exit Sub
Else
    lngRecs& = snapOvertime.RecordCount
    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
    Title$ = "Zero Overtime Bank"
    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5
Screen.MousePointer = HOURGLASS

gdbAdoIhr001.BeginTrans

While Not snapOvertime.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    EmpNo& = snapOvertime("OT_EMPNBR")
    
    DtTm = Now
        
    'If chkOvtE Then
        'If chkZeroCurrent.Value Then
            xComments = "Current Ovt. Bank Chg from " & snapOvertime("OT_BANK") & " to 0"
            Call Append_Accrual(EmpNo&, "BANK", Date, -Val(snapOvertime("OT_BANK") & ""), "Z", xComments)
            snapOvertime("OT_BANK") = 0
        'End If
        'If chkZeroPrev.Value Then
            xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to 0"
            Call Append_Accrual(EmpNo&, "BANK", Date, -Val(snapOvertime("OT_PBANK") & ""), "Z", xComments)
            snapOvertime("OT_PBANK") = 0
        'End If
        snapOvertime("OT_LDATE") = Now
        snapOvertime("OT_LTIME") = Time$
        snapOvertime("OT_LUSER") = glbLEE_ID
        snapOvertime.Update
    'End If

lblNextZRec:
    snapOvertime.MoveNext
    DoEvents
Wend

gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodType = 0
snapOvertime.Close
Screen.MousePointer = DEFAULT

Exit Sub

Overtime_Bank_ZeroOut_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Zero Out Bank", "HR_OVERTIME_BANK", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Function CR_SnapOvertime()
Dim SQLQ As String, SQLQ1 As String
Dim snapMultiEmp As New ADODB.Recordset

CR_SnapOvertime = False
On Error GoTo CR_SnapOvertime_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ

SQLQ = "SELECT OT_EMPNBR,OT_PBANK,OT_BANK,OT_BANKT,"
SQLQ = SQLQ & " OT_LUSER,OT_LDATE,OT_LTIME"
SQLQ = SQLQ & " FROM HR_OVERTIME_BANK, HREMP "
SQLQ = SQLQ & " WHERE OT_EMPNBR = ED_EMPNBR AND " & fglbESQLQ
'If Len(clpCode(4).Text) > 0 Then
'    SQLQ = SQLQ & " AND ED_EMPNBR IN "
'    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(4).Text & "') "
'End If
If snapOvertime.State <> 0 Then snapOvertime.Close
If glbOracle Then
    snapOvertime.CursorLocation = adUseServer
End If
snapOvertime.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic  'adLockPessimistic ''

CR_SnapOvertime = True

Exit Function

CR_SnapOvertime_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapOvertime", "Overtime Bank", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

'Private Sub modSetFGlobals(strTyp$)
'If strTyp$ = "Sick" Then
'    fglbSick% = True
'    fglbVac% = False
'    If glbCompEntSick$ = "M" Then
'        fglbCompMonthly% = True
'        Call modMaximums(True)
'    Else
'        fglbCompMonthly% = False
'        Call modMaximums(False)
'    End If
'    ffieldEntitle$ = "ED_SICK"
'    ffieldPEntitle$ = "ED_PSICK"
'    ffieldTEntitle$ = "ED_SICKT"
'    fglbCode$ = "SIC"
'Else
'    fglbSick% = False
'    fglbVac% = True
'    If glbCompEntVac$ = "M" Then
'        fglbCompMonthly% = True
'        Call modMaximums(True)
'    Else
'        fglbCompMonthly% = False
'        Call modMaximums(False)
'    End If
'    ffieldEntitle$ = "ED_VAC"
'    ffieldPEntitle$ = "ED_PVAC"
'    ffieldTEntitle$ = "ED_VACT"
'    fglbCode$ = "VAC"
'End If
'
'End Sub

Private Function modZeroSelection()
Dim EmpNo&
Dim dblEntitle#, dblPrevEntitle#
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, X%, y%, z%, dblNewEntitle#
Dim dblNewMax#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%
Dim xKey, xBankCodeV, xEntV, xFDateV, xTDateV, xBankCodeS, xEntS, xFDateS, xTDateS

' Entitlements are always valued in HOURS - if you enter days then it
'   works out how many hours (based on average Hrswrked/day found in salary master record)
On Error GoTo modZeroSelection_Err

modZeroSelection = False

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
'    If chkOvtE Then
'        Call Overtime_Bank_ZeroOut
'    End If
'End If

'If chkVacation Or chkSick Then

    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
              
    Screen.MousePointer = DEFAULT
    
    If snapEntitle.BOF And snapEntitle.EOF Then
        MsgBox "Employees for this selection do not exist!"
        Exit Function
    Else
        lngRecs& = snapEntitle.RecordCount
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
        Title$ = "Update Entitlements"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Function
        End If
    
    End If
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    Dim xComments
    'Ticket #11992, Don't use BeginTrans because the Integration is called in the loop
    'gdbAdoIhr001.BeginTrans
    
    While Not snapEntitle.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        
        EmpNo& = snapEntitle("ED_EMPNBR")
        
        DtTm = Now
        xBankCodeS = "": xEntS = 0: xFDateS = "": xTDateS = ""
        'If chkSick Then
            xBankCodeV = "SICK"
            xFDateV = snapEntitle("ED_EFDATES")
            xTDateV = snapEntitle("ED_ETDATES")
            'If chkZeroCurrent.Value Then
                xComments = "Current Sick Ent. Chg from " & snapEntitle("ED_SICK") & " to 0"
                'Call Append_Accrual(EmpNo&, "SICK", Date, -Val(snapEntitle("ED_SICK") & ""), "Z", xComments)
                Call Append_Accrual(EmpNo&, "SICK", snapEntitle("ED_ETDATES"), -Val(snapEntitle("ED_SICK") & ""), "Z", xComments)
                snapEntitle("ED_SICK") = 0
                snapEntitle("ED_ANNSICK") = 0
                xEntS = 0
            'End If
            'If chkZeroPrev.Value Then
                xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to 0"
                'Call Append_Accrual(EmpNo&, "SICK", Date, -Val(snapEntitle("ED_PSICK") & ""), "Z", xComments)
                Call Append_Accrual(EmpNo&, "SICK", snapEntitle("ED_ETDATES"), -Val(snapEntitle("ED_PSICK") & ""), "Z", xComments)
                snapEntitle("ED_PSICK") = 0
                xEntS = snapEntitle("ED_SICK")
            'End If
        'End If
        xBankCodeV = "": xEntV = 0: xFDateV = "": xTDateV = ""
        'If chkVacation Then
            xBankCodeV = "VAC"
            xFDateV = snapEntitle("ED_EFDATE")
            xTDateV = snapEntitle("ED_ETDATE")
            'If chkZeroCurrent.Value Then
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to 0"
                'Call Append_Accrual(EmpNo&, "VAC", Date, -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
                Call Append_Accrual(EmpNo&, "VAC", snapEntitle("ED_ETDATE"), -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
                snapEntitle("ED_VAC") = 0
                snapEntitle("ED_ANNVAC") = 0
                xEntV = 0
            'End If
            'If chkZeroPrev.Value Then
                xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to 0"
                'Call Append_Accrual(EmpNo&, "VAC", Date, -Val(snapEntitle("ED_PVAC") & ""), "Z", xComments)
                Call Append_Accrual(EmpNo&, "VAC", snapEntitle("ED_ETDATE"), -Val(snapEntitle("ED_PVAC") & ""), "Z", xComments)
                snapEntitle("ED_PVAC") = 0
                xEntV = snapEntitle("ED_VAC")
            'End If
        'End If
    
    
        snapEntitle.Update
        
        'If chkVacation Then
            xKey = EmpNo&
            xKey = xKey & "|" & Format(xFDateV, "dd-mmm-yyyy")
            xKey = xKey & "|" & Format(xTDateV, "dd-mmm-yyyy")
            xKey = xKey & "|VAC"
            xKey = xKey & "|" & xEntV
            xKey = xKey & "|" & Format(xFDateV, "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
            Call Entitlements_Master_Integration(xKey, EmpNo&) 'George added for Advance Tracker
            DoEvents
        'End If
        'If chkSick Then
            xKey = EmpNo&
            xKey = xKey & "|" & Format(xFDateS, "dd-mmm-yyyy")
            xKey = xKey & "|" & Format(xTDateS, "dd-mmm-yyyy")
            xKey = xKey & "|SICK"
            xKey = xKey & "|" & xEntS
            xKey = xKey & "|" & Format(xFDateS, "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
            Call Entitlements_Master_Integration(xKey, EmpNo&) 'George added for Advance Tracker
            DoEvents
        'End If
    
lblNextZRec:
        snapEntitle.MoveNext
        DoEvents
    
    Wend
    modZeroSelection = True
    MDIMain.panHelp(0).FloodType = 0
    snapEntitle.Close
    Screen.MousePointer = DEFAULT
    'gdbAdoIhr001.CommitTrans

'End If

Exit Function

modZeroSelection_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Roll-Over Entitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub DisplayRule(rsTA As ADODB.Recordset)
Dim SQLQ, xOrder, nOrder, aa
Dim rsVE As New ADODB.Recordset
Dim X
If Not rsTA.EOF Then
     SQLQ = "SELECT * FROM HRVACENT "
    If IsNull(rsTA("VE_DIV")) Then
        SQLQ = SQLQ & " WHERE VE_DIV IS NULL"
    Else
        SQLQ = SQLQ & " WHERE VE_DIV = '" & rsTA("VE_DIV") & "'"
    End If
    If IsNull(rsTA("VE_DEPT")) Then
        SQLQ = SQLQ & " AND VE_DEPT IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_DEPT = '" & rsTA("VE_DEPT") & "'"
    End If
    If IsNull(rsTA("VE_ORG")) Then
        SQLQ = SQLQ & " AND VE_ORG IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_ORG = '" & rsTA("VE_ORG") & "'"
    End If
    If Not IsNull(rsTA("VE_EDATE")) Then
        SQLQ = SQLQ & " AND VE_EDATE = " & Date_SQL(rsTA("VE_EDATE"))
    End If
    If IsNull(rsTA("VE_EMP")) Then
        SQLQ = SQLQ & " AND VE_EMP IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_EMP = '" & rsTA("VE_EMP") & "'"
    End If
    If IsNull(rsTA("VE_PT")) Then
        SQLQ = SQLQ & " AND VE_PT IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_PT = '" & rsTA("VE_PT") & "' "
    End If
    If IsNull(rsTA("VE_GRPCD")) Then
        SQLQ = SQLQ & " AND VE_GRPCD IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_GRPCD = '" & rsTA("VE_GRPCD") & "'"
    End If
    SQLQ = SQLQ & " Order By VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_ORDER "

    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    Do While Not rsVE.EOF
        xOrder = rsVE("VE_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 16) Then
            If Not IsNull(rsVE("VE_BMONTH")) Then medLTServ(nOrder) = rsVE("VE_BMONTH")
            If Not IsNull(rsVE("VE_EMONTH")) Then medGTServ(nOrder) = rsVE("VE_EMONTH")
            If Not IsNull(rsVE("VE_ENTITLE")) Then medEntitle(nOrder) = rsVE("VE_ENTITLE")
            If rsVE("VE_TYPE") = "D" Then optD(nOrder) = True
            If rsVE("VE_TYPE") = "H" Then optH(nOrder) = True
            If rsVE("VE_TYPE") = "F" Then optF(nOrder) = True
            If Not IsNull(rsVE("VE_MAX")) Then medMax(nOrder) = rsVE("VE_MAX")
            If Not IsNull(rsVE("VE_PCT")) Then medVacation(nOrder) = rsVE("VE_PCT")
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close

End If
rsTA.Close

End Sub


Private Sub getWSQLQ()

fglbESQLQ = glbSeleDeptUn
'If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & clpDept.Text & "' "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
'
'If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
'If glbLinamar Then
'    If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & clpCode(1).Text & "' "
'Else
'If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(1).Text & "' "
'End If
'If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
'If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
'If clpPT.Text <> "" Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
'If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

End Sub

Private Function AccuValForMulti(EmpNo, dblEnt) ' Ticket #3304
'Please talk to Jaddy
'''Dim rsJOB As New ADODB.Recordset
'''Dim xVal, SQLQ
'''xVal = 0
'''If glbMulti Then
'''    SQLQ = "SELECT JH_EMPNBR FROM qry_JobCurrent WHERE JH_EMPNBR = " & EmpNo & " "
'''    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
'''
'''Dim xVal, SQLQ
'''    xVal = 0
'''    If glbMulti Then
'''        If Len(Memplist1) > 0 Then
'''            If InStr(1, Memplist1, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist1
'''                If InStr(1, Memplist2, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist2
'''                    'xVal = 0 ' First time replace the Emtitlement with the New one
'''                    Memplist2 = Replace(Memplist2, "'" & EmpNo & "',", ",")
'''                Else
'''                    xVal = dblEnt
'''                End If
'''            End If
'''        End If
'''    End If
'''    AccuValForMulti = xVal
End Function
Private Function GetFTEtot(EmpNo, dblFTE)
'Please talk to Jaddy
'''Dim rsFTE As New ADODB.Recordset
'''Dim SQLQ, xFTE
'''    xFTE = dblFTE
'''    If glbMulti Then
'''        If Len(Memplist1) > 0 Then
'''            If InStr(1, Memplist1, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist1
'''                SQLQ = "SELECT JH_EMPNBR, SUM(JH_FTENUM) AS TOTFTE FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & EmpNo & " "
'''                SQLQ = SQLQ & "GROUP BY JH_EMPNBR "
'''                rsFTE.Open SQLQ, gdbAdoIhr001, adOpenStatic
'''                If Not rsFTE.EOF Then
'''                    If Not IsNull(rsFTE("TOTFTE")) Then
'''                        xFTE = rsFTE("TOTFTE")
'''                    End If
'''                End If
'''                rsFTE.Close
'''            End If
'''        End If
'''    End If
'''    GetFTEtot = xFTE
End Function

Private Function CalcASLRepaid(xEmpNo, xAsofDate, dblEntUpd, dblNewEnt, dblEnt#) '
Dim rsASL As New ADODB.Recordset
Dim rsENT As New ADODB.Recordset
Dim SQLQ, xTaken, xRepaid, xOutStand
Dim xSickEnt
    xSickEnt = dblEntUpd
    SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES,ED_SICK,ED_SICKT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsENT.EOF Then
        If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
            'If rsENT("ED_EFDATES") <= CVDate(xAsofDate) And rsENT("ED_ETDATES") >= CVDate(xAsofDate) Then
                SQLQ = "SELECT AS_EMPNBR, Sum(AS_HRSTAK) AS TAKEN, Sum(AS_HRSREP) AS REPAID FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                'Don't check Date Range for ASL T#3304
                'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                SQLQ = SQLQ & "GROUP BY AS_EMPNBR "
                rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                xTaken = 0: xRepaid = 0: xOutStand = 0
                If Not rsASL.EOF Then
                    If IsNull(rsASL("TAKEN")) Then
                        xTaken = 0
                    Else
                        xTaken = rsASL("TAKEN")
                    End If
                    If IsNull(rsASL("REPAID")) Then
                        xRepaid = 0
                    Else
                        xRepaid = rsASL("REPAID")
                    End If
                    xOutStand = xTaken - xRepaid
                End If
                rsASL.Close
                
                'Logic changed:
                'Repaid = Sick Entitlement, before Repaid = ASL Outstanding
                
                'xOutStand = dblEntUpd
                If xOutStand > 0 Then
                    If xOutStand >= dblNewEnt Then
                        xSickEnt = dblEnt#
                    Else
                        xSickEnt = dblEnt# + dblNewEnt - xOutStand
                        dblNewEnt = xOutStand
                    End If

                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                    'Don't check Date Range for ASL T#3304
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_DOA = ('" & Format(xAsofDate, "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_CODE = 'REPA' "
                    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    'If rsASL.EOF Then
                        rsASL.AddNew
                        rsASL("AS_HRSTAK") = 0
                    'Else
                        'dblEntUpd = dblEntUpd + rsASL("AS_HRSREP")
                    'End If
                    rsASL("AS_COMPNO") = "001"
                    rsASL("AS_EMPNBR") = xEmpNo
                    rsASL("AS_DOA") = xAsofDate
                    rsASL("AS_CODE") = "REPA"
                    rsASL("AS_HRSREP") = dblNewEnt 'dblEntUpd
                    rsASL("AS_HRSOS") = xOutStand - dblNewEnt 'dblEntUpd
                    'rsASL("AS_EFDATES") = rsENT("ED_EFDATES")
                    'rsASL("AS_ETDATES") = rsENT("ED_ETDATES")
                    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
                    rsASL("AS_LTIME") = Time$
                    rsASL("AS_LUSER") = glbUserID
                    rsASL.Update
                    rsASL.Close
                    Call ReCalcASL(xEmpNo, "")
                    'SQLQ = "UPDATE WHSCC_ASL SET AS_HRAOS = 0 "
                    'SQLQ = SQLQ & "WHERE AS_EMPNBR = " & xEmpNo & " "
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    'gdbAdoIhr001.Execute SQLQ
                End If
            'End If
        End If
    End If
    rsENT.Close
    CalcASLRepaid = xSickEnt
End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    TF = True
    UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID)
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

