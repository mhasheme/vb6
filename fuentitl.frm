VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUEntitle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Entitlements"
   ClientHeight    =   7905
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
   ScaleHeight     =   7905
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbAnnMonth 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1880
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Select Anniversary Month"
      Top             =   2700
      Width           =   765
   End
   Begin VB.CommandButton cmdZeroOutHourly 
      Appearance      =   0  'Flat
      Caption         =   "&Zero Out Hourly Entitlements"
      Height          =   375
      Left            =   4560
      TabIndex        =   164
      Tag             =   "Zero Out Hourly Entitlements"
      Top             =   7440
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.CommandButton cmdRolloverHourly 
      Appearance      =   0  'Flat
      Caption         =   "&Rollover Hourly Entitlements"
      Height          =   375
      Left            =   1320
      TabIndex        =   163
      Tag             =   "Rollover Hourly Entitlements"
      Top             =   7440
      Visible         =   0   'False
      Width           =   3105
   End
   Begin Threed.SSCheck chkOvtE 
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   3900
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Overtime Bank"
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
   Begin Threed.SSCheck chkVacation 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   3900
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Vacation Time"
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
   Begin Threed.SSCheck chkSick 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3900
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Sick Time"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Top             =   2340
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin VB.Frame VacFram02 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   840
      TabIndex        =   88
      Top             =   8520
      Visible         =   0   'False
      Width           =   7815
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   10
         Left            =   9390
         TabIndex        =   89
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
         TabIndex        =   90
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
         TabIndex        =   91
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
         TabIndex        =   92
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
         TabIndex        =   93
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
         TabIndex        =   94
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
         TabIndex        =   95
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
         TabIndex        =   96
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
         TabIndex        =   97
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
         TabIndex        =   98
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
         TabIndex        =   99
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
         TabIndex        =   100
         Top             =   0
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
               Size            =   27
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
               Size            =   27
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
               Size            =   27
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
         TabIndex        =   104
         Top             =   330
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   105
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
               Size            =   27
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
            TabIndex        =   106
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
               Size            =   27
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
            TabIndex        =   107
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
               Size            =   27
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
         TabIndex        =   108
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
         TabIndex        =   109
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
         TabIndex        =   110
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
         TabIndex        =   111
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
         TabIndex        =   112
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
         TabIndex        =   113
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
         TabIndex        =   114
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
         TabIndex        =   115
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
         TabIndex        =   116
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
         TabIndex        =   117
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
         TabIndex        =   118
         Top             =   660
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   119
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
               Size            =   27
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
            TabIndex        =   120
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
               Size            =   27
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
            TabIndex        =   121
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
         TabIndex        =   122
         Top             =   975
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   123
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
               Size            =   27
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
            TabIndex        =   124
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
               Size            =   27
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
            TabIndex        =   125
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
               Size            =   27
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
         TabIndex        =   126
         Top             =   1275
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   127
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
               Size            =   27
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
            TabIndex        =   128
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
               Size            =   27
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
            TabIndex        =   129
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
               Size            =   27
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
         TabIndex        =   130
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
         TabIndex        =   131
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
         TabIndex        =   132
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
         TabIndex        =   133
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
         TabIndex        =   134
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
         TabIndex        =   135
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
         TabIndex        =   136
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
         TabIndex        =   137
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
         TabIndex        =   138
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
         TabIndex        =   139
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
         TabIndex        =   140
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
         TabIndex        =   141
         Top             =   1600
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   142
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
               Size            =   27
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
            TabIndex        =   143
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
               Size            =   27
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
            TabIndex        =   144
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
               Size            =   27
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
         TabIndex        =   145
         Top             =   1940
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
            TabIndex        =   146
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
               Size            =   27
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
            TabIndex        =   147
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
               Size            =   27
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
            TabIndex        =   148
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
               Size            =   27
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
         TabIndex        =   149
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
         TabIndex        =   150
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
         TabIndex        =   151
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
         Top             =   2040
         Width           =   1440
      End
   End
   Begin Threed.SSFrame frmZero 
      Height          =   540
      Left            =   0
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   9120
      _Version        =   65536
      _ExtentX        =   16087
      _ExtentY        =   952
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSCheck chkZeroCurrent 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   195
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Current Year"
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
      Begin Threed.SSCheck chkZeroPrev 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   195
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Previous Year"
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
   End
   Begin Threed.SSFrame frmRoll 
      Height          =   795
      Left            =   20
      TabIndex        =   27
      Top             =   6000
      Visible         =   0   'False
      Width           =   9120
      _Version        =   65536
      _ExtentX        =   16087
      _ExtentY        =   1402
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optReplace 
         Height          =   225
         Left            =   4440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   3945
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Outstanding Balances"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optAccumulate 
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Negative Balances Only"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox medLTServ 
      Height          =   285
      Index           =   4
      Left            =   2820
      TabIndex        =   28
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
      TabIndex        =   29
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
      TabIndex        =   30
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
      TabIndex        =   31
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
      TabIndex        =   32
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
      TabIndex        =   33
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
      TabIndex        =   34
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
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   37
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
      TabIndex        =   38
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
      TabIndex        =   39
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
      TabIndex        =   40
      Top             =   9180
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
         TabIndex        =   41
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
            Size            =   27
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
         TabIndex        =   42
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
            Size            =   27
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
         TabIndex        =   43
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
      TabIndex        =   44
      Top             =   9540
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
         TabIndex        =   45
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
            Size            =   27
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
         TabIndex        =   46
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
            Size            =   27
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
         TabIndex        =   47
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
            Size            =   27
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
      TabIndex        =   48
      Top             =   9870
      Visible         =   0   'False
      Width           =   2640
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
            Size            =   27
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
         TabIndex        =   50
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
            Size            =   27
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
         TabIndex        =   51
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
            Size            =   27
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
      TabIndex        =   52
      Top             =   10200
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
            Size            =   27
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
            Size            =   27
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
            Size            =   27
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
      TabIndex        =   56
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
      TabIndex        =   57
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
      TabIndex        =   58
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
      TabIndex        =   59
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
      TabIndex        =   60
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
      TabIndex        =   61
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
      TabIndex        =   62
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
      TabIndex        =   63
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
      TabIndex        =   68
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
      TabIndex        =   69
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
      TabIndex        =   70
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
      TabIndex        =   71
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
      TabIndex        =   72
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
      TabIndex        =   73
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
      TabIndex        =   74
      Top             =   10590
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
            Size            =   27
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
         TabIndex        =   76
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
            Size            =   27
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
         TabIndex        =   77
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
      TabIndex        =   78
      Top             =   10950
      Visible         =   0   'False
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
            Size            =   27
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
            Size            =   27
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
            Size            =   27
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
      TabIndex        =   82
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
      TabIndex        =   83
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
      TabIndex        =   84
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
      TabIndex        =   85
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Status Code"
      Top             =   1350
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Union Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "10-Enter Employee Number"
      Top             =   2010
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   5
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   360
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   6
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   690
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin Threed.SSOption optDH 
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   167
      TabStop         =   0   'False
      Tag             =   "Hours to Rollover"
      Top             =   4290
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Hours"
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
   Begin MSMask.MaskEdBox medMaxRollover 
      Height          =   285
      Left            =   3360
      TabIndex        =   169
      Tag             =   "10-Maximum Hours/Days to Rollover"
      Top             =   4245
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optDH 
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   168
      Tag             =   "Days to Rollover"
      Top             =   4290
      Width           =   690
      _Version        =   65536
      _ExtentX        =   1217
      _ExtentY        =   344
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
   End
   Begin VB.Label lblMaxRollover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Hours or Days to Rollover"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   166
      Top             =   4290
      Width           =   3135
   End
   Begin VB.Label lblAnnMonth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Anniversary Month"
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
      TabIndex        =   165
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   5280
      TabIndex        =   162
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      Left            =   5280
      TabIndex        =   161
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   160
      Top             =   1710
      Width           =   630
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      TabIndex        =   159
      Top             =   2070
      Width           =   1290
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
      TabIndex        =   87
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
      TabIndex        =   86
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      TabIndex        =   65
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
      TabIndex        =   64
      Top             =   10320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblRollEntitlements 
      Caption         =   "Rollover Entitlements"
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
      Left            =   8760
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblZero 
      Caption         =   "Zero Out Entitlements"
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
      Left            =   8760
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   1260
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
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label textMulti 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The Union Code and FT/PT/SE/TR/OT will be validated from the Employee Basic Data"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3210
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      TabIndex        =   20
      Top             =   1350
      Width           =   1350
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
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
      TabIndex        =   19
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      TabIndex        =   18
      Top             =   690
      Width           =   825
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
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmUEntitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim snapEntitle As New adodb.Recordset     'user vier
Dim snapOvertime As New adodb.Recordset
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

'Private Sub cmdEntitl_GotFocus()
'Call SetPanHelp(ActiveControl)

'End Sub

Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%

On Error GoTo Mod_Err
If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
If (Not chkSick.Value) And (Not chkVacation.Value) And (Not chkOvtE.Value) Then
    MsgBox "You must select at least one type of update (Sick Time, Vacation Time and Overtime Bank)"
    Exit Sub
End If
If (Not chkZeroCurrent.Value) And (Not chkZeroPrev.Value) And frmZero.Visible = True Then
    MsgBox "You must select at least one type of update (Current Year, Previous Year)"
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

Select Case fglbModifyWhat$
    Case "Zero Out"
        If Not modZeroSelection() Then Exit Sub
    Case "Rollover"
        If Len(medMaxRollover) > 0 Then
            If Not IsNumeric(medMaxRollover) Then
                Screen.MousePointer = DEFAULT
                MsgBox "Invalid Maximum Hours/Days to Rollover."
                Exit Sub
            End If
        End If
        '7.9 - Ticket #20020 - Jerry asked to do the Recalculate before the rollover
        Call EntReCalc(fglbESQLQ, Empty, "TAKEN ONLY")
        If Not modRollSelection() Then Exit Sub
End Select

Call EntReCalc(fglbESQLQ)
Screen.MousePointer = DEFAULT

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

Public Sub cmdRollover_Click()
'panJGroup.Visible = False       'js-6Apr99


ControlsShown = False
Call UpdateEntControls(ControlsShown)
fglbModifyWhat$ = "Rollover"

frmZero.Visible = False '
frmRoll.Visible = True
frmRoll.Top = 4600
lblRollEntitlements.Visible = True
lblRollEntitlements.Left = 120
lblRollEntitlements.Top = 3600
Me.Caption = "Rollover Entitlements"

medMaxRollover.Text = ""
lblMaxRollover.Visible = True
medMaxRollover.Visible = True
optDH(0).Visible = True
optDH(1).Visible = True

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then

    chkOvtE.Visible = True
    If glbCompSerial = "S/N - 2418W" Then
        chkVacation.SetFocus
        chkVacation.Value = True
    ElseIf glbCompSerial = "S/N - 2430W" Then  'Ticket #27729 Franks 03/14/2016 Carizon Rollover only
        chkSick = False
        chkSick.Enabled = False
        chkVacation.SetFocus
        chkVacation.Value = True
    Else
        chkSick.SetFocus
        chkSick.Value = True
    End If
    
'Else
'    optOvtE.Visible = False
'End If

End Sub

'Private Sub cmdRollover_GotFocus()
'Call SetPanHelp(ActiveControl)

'End Sub

Public Sub cmdZeroOut_Click()
'optBothVS.Visible = True

fglbModifyWhat$ = "Zero Out"
ControlsShown = False
Call UpdateEntControls(ControlsShown)

lblRollEntitlements.Visible = False
lblZero.Visible = True 'js-6Apr99
lblZero.Left = 120     '
lblZero.Top = 3600     '

frmZero.Visible = True '
frmZero.Top = 4600

frmRoll.Visible = False
Me.Caption = "Zero Out Entitlements"

medMaxRollover.Text = ""
lblMaxRollover.Visible = False
medMaxRollover.Visible = False
optDH(0).Visible = False
optDH(1).Visible = False

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    chkOvtE.Visible = True
    If glbCompSerial = "S/N - 2418W" Then
        chkVacation.SetFocus
        chkVacation.Value = True
    Else
        chkSick.SetFocus
        chkSick.Value = True
    End If
'Else
'    optOvtE.Visible = False
'End If

End Sub

'Private Sub cmdZeroOut_GotFocus()
'Call SetPanHelp(ActiveControl)

'End Sub

Private Function CR_SnapEntitle()
Dim SQLQ As String, SQLQ1 As String
Dim snapMultiEmp As New adodb.Recordset

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
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
SQLQ = SQLQ & " ED_OTBANK, "
End If
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME"
SQLQ = SQLQ & " FROM HREMP "
SQLQ = SQLQ & " WHERE " & fglbESQLQ
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(4).Text & "') "
End If
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

Private Sub cmbAnnMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comAnnMonthAdding()
'When selected by the users, the report will only show employees who have their Original Date of Hire month
'equal the Anniversary Month.
    cmbAnnMonth.AddItem ""
    cmbAnnMonth.AddItem "Jan"
    cmbAnnMonth.AddItem "Feb"
    cmbAnnMonth.AddItem "Mar"
    cmbAnnMonth.AddItem "Apr"
    cmbAnnMonth.AddItem "May"
    cmbAnnMonth.AddItem "Jun"
    cmbAnnMonth.AddItem "Jul"
    cmbAnnMonth.AddItem "Aug"
    cmbAnnMonth.AddItem "Sep"
    cmbAnnMonth.AddItem "Oct"
    cmbAnnMonth.AddItem "Nov"
    cmbAnnMonth.AddItem "Dec"
End Sub

Private Sub cmdRolloverHourly_Click()
    'Ticket #17924 - Begin
    Unload frmUHrsEnt
    Load frmUHrsEnt
    frmUHrsEnt.Caption = "Rollover Hourly Entitlement"
    frmUHrsEnt.cmdRolloverHr_Click
    frmUHrsEnt.ZOrder 0
    'Ticket #17924 - End
    
End Sub

Private Sub cmdZeroOutHourly_Click()
    'Ticket #17924 - Begin
    Unload frmUHrsEnt
    Load frmUHrsEnt
    frmUHrsEnt.Caption = "Zero Out Hourly Entitlement"
    frmUHrsEnt.cmdZeroOutHr_Click
    frmUHrsEnt.ZOrder 0
    'Ticket #17924 - End
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUENTITLE"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUENTITLE"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim rsVE As New adodb.Recordset
If glbCompSerial = "S/N - 2418W" Then 'ticket# 17786
    chkSick.Visible = False
    chkSick = False
    chkVacation = True
Else
    chkSick = True  'default to sick
End If

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)
Call comAnnMonthAdding

UpdVac = False

Call modSetFGlobals("Sick")

If glbMulti Then textMulti.Visible = True
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

If glbLinamar Then
    lblSection = "Vacation Group"
    clpCode(1).LookupType = SalaryDistribution
    
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

If glbUEnt = 2 Then
    Me.Caption = "Rollover Entitlements"
    
    'Ticket #17924 - Begin
    'cmdRolloverHourly.Visible = True   'Menu Item added
    'cmdZeroOutHourly.Visible = False   'Menu Item added
    'Ticket #17924 - End
Else
    Me.Caption = "Zero Out Entitlements"
        
    'Ticket #17924 - Begin
    'cmdRolloverHourly.Visible = False  'Menu Item added
    'cmdZeroOutHourly.Visible = True    'Menu Item added
    'Ticket #17924 - End
End If

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
Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Sub medEntitle_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medGTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medLTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMax_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medVacation_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
If IsNumeric(medVacation(Index)) Then
    If Len(medVacation(Index)) > 0 Then
        medVacation(Index) = medVacation(Index) * 100
    End If
End If

End Sub

Private Sub medVacation_LostFocus(Index As Integer)
'If (Not IsNumeric(medVacation(Index))) Then medVacation(Index) = 0
If IsNumeric(medVacation(Index)) Then
    If Len(medVacation(Index)) > 0 Then
        medVacation(Index) = medVacation(Index) / 100
    End If
End If
End Sub

Private Sub modMaximums(TF%)
Dim x%

End Sub

Private Function modRollSelection()
Dim empNo As Long, strJob$, spt As Variant, lngRecs&
Dim x%, DtTm As Variant, Msg$, Title$, DgDef As Variant
Dim Response%, pct%, prec%, xErr
Dim SQLQ As String, dblOUTS#, dblOUTV#
Dim xHrsDay
Dim xSkipped As String
Dim flgSkip As Boolean

On Error GoTo modRollSelection_Err
modRollSelection = False

xErr = False
Msg$ = ""
If (chkSick) And (glbEntOutStandingS$ <> "1" And Len(cmbAnnMonth) = 0) Then     'Ticket #18721 - Anniversary Month
    Msg$ = Msg$ & "SickTime Outstanding Entitlements" & Chr(10) & "is not based on Entitlement Date" & Chr(10)
    xErr = True
End If
If (chkVacation) And (glbEntOutStanding$ <> "1" And Len(cmbAnnMonth) = 0) Then  'Ticket #18721 - Anniversary Month
    Msg$ = Msg$ & "Vacation Outstanding Entitlements" & Chr(10) & "is not based on Entitlement Date" & Chr(10)
    xErr = True
End If
If xErr Then
    Title$ = "RollOver Aborted"
    Msg$ = Msg$ & "Rollover cannot continue !!!"
    DgDef = MB_ICONEXCLAMATION
    MsgBox Msg$, DgDef, Title$
    Screen.MousePointer = DEFAULT
    Exit Function
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    If chkOvtE Then
        Call Overtime_Bank_Rollover
    End If
'End If

If chkVacation Or chkSick Then

    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
    
    Screen.MousePointer = DEFAULT
    xSkipped = ""
    
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
    
    'gdbAdoIhr001.BeginTrans    'Finding this is causing an issue when the Vadim Integration is On. The rollover does not happen but does not give any error.
    
    While Not snapEntitle.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        empNo& = snapEntitle("ED_EMPNBR")
        
        flgSkip = False
        dblOUTS# = 0
        If chkSick Then
            If IsNumeric(snapEntitle("ED_PSICK")) Then
                  dblOUTS# = dblOUTS# + snapEntitle("ED_PSICK")
            End If
            If IsNumeric(snapEntitle("ED_SICK")) Then
                  dblOUTS# = dblOUTS# + snapEntitle("ED_SICK")
            End If
            If IsNumeric(snapEntitle("ED_SICKT")) Then
                  dblOUTS# = dblOUTS# - snapEntitle("ED_SICKT")
            End If
            
            'Frank 10/21/03 ticket #2292
            'If Dept. = 2 or 11 and Union = 2, Maximum previous year (rollover) can be 16 hours.
            If glbCElgin Then
                If (snapEntitle("ED_DEPTNO") = "2" Or snapEntitle("ED_DEPTNO") = "11") And snapEntitle("ED_ORG") = "2" Then
                    If dblOUTS# > 16 Then dblOUTS# = 16
                End If
            Else
                'Maximum Hours/Days to Rollover
                If IsNumeric(medMaxRollover) Then
                    If optDH(0) Then    'Hours
                        If Val(medMaxRollover) < dblOUTS# Then
                            dblOUTS# = medMaxRollover
                        End If
                    Else
                        'Convert Days into Hours before comparison
                        xHrsDay = GetJHData(empNo&, "JH_DHRS", 0)
                        If xHrsDay = 0 Then
                            xSkipped = xSkipped & ", " & empNo&
                            flgSkip = True
                        Else
                            If (Val(medMaxRollover) * Val(xHrsDay)) < Val(dblOUTS#) Then
                                dblOUTS# = Val(medMaxRollover) * Val(xHrsDay)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        dblOUTV# = 0
        If chkVacation Then
            If IsNumeric(snapEntitle("ED_PVAC")) Then
                  dblOUTV# = dblOUTV# + snapEntitle("ED_PVAC")
            End If
            If IsNumeric(snapEntitle("ED_VAC")) Then
                  dblOUTV# = dblOUTV# + snapEntitle("ED_VAC")
            End If
            If IsNumeric(snapEntitle("ED_VACT")) Then
                  dblOUTV# = dblOUTV# - snapEntitle("ED_VACT")
            End If
            
            'Maximum Hours/Days to Rollover
            If IsNumeric(medMaxRollover) Then
                If optDH(0) Then    'Hours
                    If Val(medMaxRollover) < Val(dblOUTV#) Then
                        dblOUTV# = medMaxRollover
                    End If
                Else
                    'Convert Days into Hours before comparison
                    xHrsDay = GetJHData(empNo&, "JH_DHRS", 0)
                    If xHrsDay = 0 Then
                        xSkipped = xSkipped & ", " & empNo&
                        flgSkip = True
                    Else
                        If (Val(medMaxRollover) * Val(xHrsDay)) < Val(dblOUTV#) Then
                            dblOUTV# = Val(medMaxRollover) * Val(xHrsDay)
                        End If
                    End If
                End If
            End If
        End If
    
        If chkSick = True And flgSkip = False Then
            If optAccumulate Then   'Negative Balances only
                If dblOUTS# < 0 Then
                    xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & dblOUTS#
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "SICK", Date, dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    'Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PSICK") = dblOUTS#
                Else
                    xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & 0
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "SICK", Date, 0 - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    'Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                    Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), 0 - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PSICK") = 0
                End If
            Else
                'Ticket #23141 - For Vadim clients Rolling over differently.
                'I will have to clear the balance in Vadim first, i.e. pass -ve OS Bal, so it becomes 0 balance in Vadim
                'and then pass OS to add back the OS. This will show the clear in and out in Accrual file and in Vadim.
                If glbVadim Then
                    'Clear the Previous from Vadim first
                    'xComments = "Vadim only: Prev. Sick Ent. Chg " & " to 0" '& dblOUTV#
                    xComments = "Vadim OS: Prev. Sick Ent. Chg from " & dblOUTS# & " to 0" '& dblOUTS#
                    Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), 0 - dblOUTS#, "R", xComments)
                End If
                
                If glbVadim Then
                    'Ticket #23141 - For Vadim it is actually changing from 0 to OS amount
                    xComments = "Prev. Sick Ent. Chg from 0" & " to " & dblOUTS#
                Else
                    xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to " & dblOUTS#
                End If
                '================= By Hemu
                'Call Append_Accrual(EmpNo&, "SICK", Date, dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                'Call Append_Accrual(EmpNo&, "SICK", Date, Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                'Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), Val(snapEntitle("ED_SICK") & ""), "R", xComments)
                If glbVadim Then
                    'Ticket #23141 - Add full OS back after clearing above
                    Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), dblOUTS#, "R", xComments)
                Else
                    Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), dblOUTS# - Val(snapEntitle("ED_PSICK") & ""), "R", xComments)
                End If
                '================= By Hemu
                
                snapEntitle("ED_PSICK") = dblOUTS#
                
                'We will have to clear the Current because there is no Zero Out for Vadim clients when doing
                'Year End as they go with the OS. Also if it's Monthly accumulation of entitlements in info:HR,
                'the new year should start with 0 current otherwise it will add to the Current. Not passing the
                'zero out to Vadim for this.
                If glbVadim Then
                    snapEntitle("ED_SICK") = 0
                    snapEntitle("ED_ANNSICK") = 0
                End If
            End If
            
            'Release 8.0 - Ticket #22682: Function to delete all Previous Year's Exceeding Follow Up records.
            If Not IsNull(snapEntitle("ED_ETDATES")) Then
                Call Delete_Exceeding_FollowUp(empNo&, "SICK", Year(snapEntitle("ED_ETDATES")))
            End If
        End If
        
        If chkVacation And flgSkip = False Then
            If optAccumulate Then
                If dblOUTV# < 0 Then
                    xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & dblOUTV#
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "VAC", Date, dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    'Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PVAC") = dblOUTV#
                Else
                    xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & 0
                    '================= By Hemu
                    'Call Append_Accrual(EmpNo&, "VAC", Date, 0 - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    'Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), 0 - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                    '================= By Hemu
                    snapEntitle("ED_PVAC") = 0
                End If
            Else
                
                'Ticket #25944 - United Way of Lower Mainland
                'Don't rollover the entitlement if the Entitlement Period is not rolling over....
                If glbCompSerial = "S/N - 2424W" Then
                    'Employee's entitlement period rolling over to new year?
                    If Not IsNull(snapEntitle("ED_ETDATE")) Then
                        If Not IsNull(snapEntitle("ED_EFDATE")) Then
                            If Len(cmbAnnMonth) > 0 Then
                                'If the Month is same and the Year is less than this year then rollover
                                If month(snapEntitle("ED_EFDATE")) = cmbAnnMonth.ListIndex And Year(snapEntitle("ED_EFDATE")) < Year(Now) Then
                                    'Do allow rolling over
                                Else
                                    'Skip the entitlement rollover
                                    GoTo Next_Step
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Ticket #23141 - For Vadim clients Rolling over differently.
                'I will have to clear the balance in Vadim first, i.e. pass -ve OS Bal, so it becomes 0 balance in Vadim
                'and then pass OS to add back the OS. This will show the clear in and out in Accrual file and in Vadim.
                If glbVadim Then
                    'Clear the Previous from Vadim first
                    'xComments = "Vadim only: Prev. Vac. Ent. Chg " & " to 0" '& dblOUTV#
                    xComments = "Vadim OS. Prev. Vac. Ent. Chg from " & dblOUTV# & " to 0" '& dblOUTV#
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), 0 - dblOUTV#, "R", xComments)
                End If
                                
                If glbVadim Then
                    'Ticket #23141 - For Vadim it is actually changing from 0 to OS amount
                    xComments = "Prev. Vac. Ent. Chg from 0" & " to " & dblOUTV#
                Else
                    xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & dblOUTV#
                End If
                '================= By Hemu
                'Call Append_Accrual(EmpNo&, "VAC", Date, dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                'Call Append_Accrual(EmpNo&, "VAC", Date, Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                'Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), Val(snapEntitle("ED_VAC") & ""), "R", xComments)
                If glbVadim Then
                    'Ticket #23141 - Add full OS back after clearing above
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), dblOUTV#, "R", xComments)
                Else
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
                End If
                '================= By Hemu
                
                snapEntitle("ED_PVAC") = dblOUTV#
            
                'We will have to clear the Current because there is no Zero Out for Vadim clients when doing
                'Year End as they go with the OS. Also if it's Monthly accumulation of entitlements in info:HR,
                'the new year should start with 0 current otherwise it will add to the Current. Not passing the
                'zero out to Vadim for this.
                If glbVadim Then
                    snapEntitle("ED_VAC") = 0
                    snapEntitle("ED_ANNVAC") = 0
                End If
            End If
                        
Next_Step:
            'Release 8.0 - Ticket #22682: Function to delete all Previous Year's Exceeding Follow Up records.
            If Not IsNull(snapEntitle("ED_ETDATE")) Then
                Call Delete_Exceeding_FollowUp(empNo&, "VAC", Year(snapEntitle("ED_ETDATE")))
            End If
            
            'Ticket #25300: United Way of Lower Mainland - not going to compute new Entitlement Period for them because
            'it will be compute when doing the Rollover with Anniversary Month.
            If glbCompSerial = "S/N - 2424W" Then
                'Update employees who fall under the Anniversary Month only to a new Entitlement Period
                If Not IsNull(snapEntitle("ED_ETDATE")) Then
                    If Not IsNull(snapEntitle("ED_EFDATE")) Then
                        If Len(cmbAnnMonth) > 0 Then
                            'If the Month is same and the Year is less than this year then rollover
                            If month(snapEntitle("ED_EFDATE")) = cmbAnnMonth.ListIndex And Year(snapEntitle("ED_EFDATE")) < Year(Now) Then
                                snapEntitle("ED_EFDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("d", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
                                snapEntitle("ED_ETDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("yyyy", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
                            End If
                        Else
                            snapEntitle("ED_EFDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("d", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
                            snapEntitle("ED_ETDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("yyyy", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
                        End If
                    End If
                End If
            End If
        End If
        
        snapEntitle("ED_LDATE") = Now
        snapEntitle("ED_LTIME") = Time$
        snapEntitle("ED_LUSER") = glbLEE_ID
        snapEntitle.Update
        snapEntitle.MoveNext
    Wend
    'gdbAdoIhr001.CommitTrans
    
    modRollSelection = True
    MDIMain.panHelp(0).FloodType = 0
    snapEntitle.Close
    Screen.MousePointer = DEFAULT
    
    If Len(xSkipped) > 0 Then
        MsgBox "Employee(s) skipped due to missing Hours/Day on Position screen for Maximum Rollover: " & xSkipped, vbExclamation, "Skipped Rollover"
    End If
End If

'Hemu - 01/14/2004 Begin - Ticket #5371
'Msg$ = "Don't forget to go into the Company" & Chr(10)
'Msg$ = Msg$ & "Master file to EDIT Entitlement Date" & Chr(10)
'Msg$ = Msg$ & "Range PRIOR to MASS UPDATING and" & Chr(10)
'Msg$ = Msg$ & "RECALCULATING Next Year's Entitlements"
'Title$ = "RollOver Update Completed"
'DgDef = MB_ICONEXCLAMATION
'Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
'Hemu - 01/14/2004 End

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
Dim RsAttHis As New adodb.Recordset
Dim rsTabl As New adodb.Recordset
Dim empNo As Long, strJob$, spt As Variant, lngRecs&
Dim x%, DtTm As Variant, Msg$, Title$, DgDef As Variant
Dim Response%, pct%, prec%, xErr
Dim SQLQ As String, dblOUTO
Dim xComments
Dim xSkipped As String
Dim xHrsDay
Dim xOrigPBANK

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
xSkipped = ""

gdbAdoIhr001.BeginTrans

While Not snapOvertime.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    empNo& = snapOvertime("OT_EMPNBR")
    
    dblOUTO = 0
    xOrigPBANK = 0
    
    If chkOvtE Then
        If IsNumeric(snapOvertime("OT_PBANK")) Then
              dblOUTO = dblOUTO + snapOvertime("OT_PBANK")
        End If
        If IsNumeric(snapOvertime("OT_BANK")) Then
              dblOUTO = dblOUTO + snapOvertime("OT_BANK")
        End If
        If IsNumeric(snapOvertime("OT_BANKT")) Then
              dblOUTO = dblOUTO - snapOvertime("OT_BANKT")
        End If
    
        'Maximum Hours/Days to Rollover
        If IsNumeric(medMaxRollover) Then
            If optDH(0) Then    'Hours
                If Val(medMaxRollover) < Val(dblOUTO) Then
                    dblOUTO = medMaxRollover
                End If
            Else
                'Convert Days into Hours before comparison
                xHrsDay = GetJHData(empNo&, "JH_DHRS", 0)
                If xHrsDay = 0 Then
                    xSkipped = xSkipped & ", " & empNo&
                    GoTo Skip_Overtime
                Else
                    If (Val(medMaxRollover) * Val(xHrsDay)) < Val(dblOUTO) Then
                        dblOUTO = Val(medMaxRollover) * Val(xHrsDay)
                    End If
                End If
            End If
        End If
    
        If optAccumulate Then
            If dblOUTO < 0 Then
                xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & dblOUTO
                Call Append_Accrual(empNo&, "BANK", Date, dblOUTO - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
                snapOvertime("OT_PBANK") = dblOUTO
            Else
                xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & 0
                Call Append_Accrual(empNo&, "BANK", Date, 0 - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
                snapOvertime("OT_PBANK") = 0
            End If
        Else
            xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to " & dblOUTO
            Call Append_Accrual(empNo&, "BANK", Date, dblOUTO - Val(snapOvertime("OT_PBANK") & ""), "R", xComments)
            snapOvertime("OT_PBANK") = dblOUTO
        End If
        
        If dblOUTO <> 0 Then
            'Add OTBF record in Attendance
            Set rsTabl = Nothing
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'OTBF' "
            rsTabl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsTabl.EOF Then
                rsTabl.AddNew
                rsTabl("TB_COMPNO") = "001"
                rsTabl("TB_NAME") = "ADRE"
                rsTabl("TB_KEY") = "OTBF"
                rsTabl("TB_DESC") = "BRING FORWARD COMP HOURS"
                rsTabl("TB_LDATE") = Date
                rsTabl("TB_LTIME") = Time$
                rsTabl("TB_LUSER") = glbUserID
                rsTabl.Update
            End If
            rsTabl.Close
            
            Set RsAttHis = Nothing
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & empNo& & " "
            SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(DateAdd("d", 1, CVDate(snapOvertime("OT_ETDATE"))))
            SQLQ = SQLQ & "AND AD_REASON = 'OTBF' "
            RsAttHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If RsAttHis.EOF Then
                RsAttHis.AddNew
            End If
            RsAttHis("AD_COMPNO") = "001"
            RsAttHis("AD_EMPNBR") = empNo&
            RsAttHis("AD_DOA") = DateAdd("d", 1, CVDate(snapOvertime("OT_ETDATE")))
            RsAttHis("AD_REASON") = "OTBF"
            RsAttHis("AD_HRS") = dblOUTO
            RsAttHis("AD_SEN") = 0 '-1 As Linda request, turn off the seniority flag
            RsAttHis("AD_LDATE") = Date
            RsAttHis("AD_LUSER") = glbUserID
            RsAttHis("AD_LTIME") = Time$
            RsAttHis.Update
            RsAttHis.Close
        End If
        
    End If
    snapOvertime("OT_LDATE") = Now
    snapOvertime("OT_LTIME") = Time$
    snapOvertime("OT_LUSER") = glbLEE_ID
    snapOvertime.Update
    
Skip_Overtime:
    snapOvertime.MoveNext
Wend
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodType = 0
snapOvertime.Close
Screen.MousePointer = DEFAULT

If Len(xSkipped) > 0 Then
    MsgBox "Employee(s) skipped due to missing Hours/Day on Position screen for Maximum Rollover: " & xSkipped, vbExclamation, "Skipped Rollover"
End If

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
Dim empNo As Long, strJob$, spt As Variant, lngRecs&
Dim x%, DtTm As Variant, Msg$, Title$, DgDef As Variant
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
    empNo& = snapOvertime("OT_EMPNBR")
    
    DtTm = Now
        
    If chkOvtE Then
        If chkZeroCurrent.Value Then
            xComments = "Current Ovt. Bank Chg from " & snapOvertime("OT_BANK") & " to 0"
            Call Append_Accrual(empNo&, "BANK", Date, -Val(snapOvertime("OT_BANK") & ""), "Z", xComments)
            snapOvertime("OT_BANK") = 0
        End If
        If chkZeroPrev.Value Then
            xComments = "Prev. Ovt. Bank Chg from " & snapOvertime("OT_PBANK") & " to 0"
            Call Append_Accrual(empNo&, "BANK", Date, -Val(snapOvertime("OT_PBANK") & ""), "Z", xComments)
            snapOvertime("OT_PBANK") = 0
        End If
        snapOvertime("OT_LDATE") = Now
        snapOvertime("OT_LTIME") = Time$
        snapOvertime("OT_LUSER") = glbLEE_ID
        snapOvertime.Update
    End If

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
Dim snapMultiEmp As New adodb.Recordset

CR_SnapOvertime = False
On Error GoTo CR_SnapOvertime_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ

SQLQ = "SELECT OT_EMPNBR,OT_PBANK,OT_BANK,OT_BANKT,"
SQLQ = SQLQ & " OT_LUSER,OT_LDATE,OT_LTIME,OT_ETDATE "
SQLQ = SQLQ & " FROM HR_OVERTIME_BANK, HREMP "
SQLQ = SQLQ & " WHERE OT_EMPNBR = ED_EMPNBR AND " & fglbESQLQ
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(4).Text & "') "
End If
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

Private Sub modSetFGlobals(strTyp$)
If strTyp$ = "Sick" Then
    fglbSick% = True
    fglbVac% = False
    If glbCompEntSick$ = "M" Then
        fglbCompMonthly% = True
        Call modMaximums(True)
    Else
        fglbCompMonthly% = False
        Call modMaximums(False)
    End If
    ffieldEntitle$ = "ED_SICK"
    ffieldPEntitle$ = "ED_PSICK"
    ffieldTEntitle$ = "ED_SICKT"
    fglbCode$ = "SIC"
Else
    fglbSick% = False
    fglbVac% = True
    If glbCompEntVac$ = "M" Then
        fglbCompMonthly% = True
        Call modMaximums(True)
    Else
        fglbCompMonthly% = False
        Call modMaximums(False)
    End If
    ffieldEntitle$ = "ED_VAC"
    ffieldPEntitle$ = "ED_PVAC"
    ffieldTEntitle$ = "ED_VACT"
    fglbCode$ = "VAC"
End If

End Sub
''''this is not useful
''''Private Function modUpdateSelection()
''''Dim EmpNo As Long
''''Dim dblEntitle#, dblPrevEntitle#, dblEntitleTaken#
''''Dim strJob$, dblServiceYears#
''''Dim spt As Variant, varStartDate As Variant, lngRecs&
''''Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
''''Dim dblFTEHours#
''''Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
''''Dim Msg$, Title$, DgDef As Variant
''''Dim Response%, pct%
''''Dim prec%
''''Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
''''Dim if_Entitle As Boolean, if_Vacation As Boolean
''''
''''On Error GoTo modUpdateSelection_Err
''''modUpdateSelection = False
''''
''''
''''If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
''''Screen.MousePointer = Default
''''If snapEntitle.BOF And snapEntitle.EOF Then
''''    MsgBox "Employees for this selection do not exist!"
''''    Exit Function
''''Else
''''    lngRecs& = snapEntitle.RecordCount
''''    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
''''    Title$ = "Update Entitlements"
''''    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
''''    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
''''    If Response% = IDNO Then    ' Evaluate response
''''        Exit Function
''''    End If
''''    Screen.MousePointer = HOURGLASS
''''End If
''''MDIMain.panHelp(0).FloodType = 1
''''MDIMain.panHelp(0).FloodPercent = 5
''''
''''For X% = 0 To 16
''''    If Not IsNumeric(medLTServ(X%)) Then Exit For ' medLTServ(X%) = 0
''''    If Not IsNumeric(medGTServ(X%)) Then
''''      medGTServ(X%) = 0
''''    Else
''''      If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
''''    End If
''''    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
''''Next
''''
''''gdbAdoIhr001.BeginTrans
''''
''''While Not snapEntitle.EOF
''''    prec% = prec% + 1
''''    pct% = Int(100 * (prec% / lngRecs&))
''''    MDIMain.panHelp(0).FloodPercent = pct%
''''    if_Entitle = False
''''    if_Vacation = False
''''
''''    EmpNo& = snapEntitle("ED_EMPNBR")
''''
''''    If IsNull(snapEntitle(ffieldEntitle$)) Then
''''        dblEntitle# = 0
''''    Else
''''        dblEntitle# = snapEntitle(ffieldEntitle$)
''''    End If
''''
''''    If IsNull(snapEntitle(ffieldPEntitle$)) Then
''''        dblPrevEntitle# = 0
''''    Else
''''        dblPrevEntitle# = snapEntitle(ffieldPEntitle$)
''''    End If
''''
''''    'Frank #3646 For Simcoe County Health Unit
''''    If IsNull(snapEntitle(ffieldTEntitle$)) Then
''''        dblEntitleTaken# = 0
''''    Else
''''        dblEntitleTaken# = snapEntitle(ffieldTEntitle$)
''''    End If
''''    'Frank #3646 For Simcoe County Health Unit
''''
''''    spt = snapEntitle("ED_PT")
''''
''''    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
''''
''''    varStartDate = snapEntitle(fglbWDate$)
''''
''''    Dim rsJOB As New ADODB.Recordset
''''    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
''''    dblDHours# = 0
''''    dblFTEHours# = 0
''''    'This loop is for both of multi positions and not multi
''''    Do Until rsJOB.EOF
''''        If IsNumeric(rsJOB("JH_DHRS")) Then
''''            dblDHours# = dblDHours# + rsJOB("JH_DHRS")
''''            If IsNumeric(rsJOB("JH_FTENUM")) Then
''''                dblFTEHours# = dblFTEHours# + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
''''                'FTE HOURS = FTE NUMBER * HOURS PER DAY
''''            End If
''''        End If
''''        rsJOB.MoveNext
''''    Loop
''''    rsJOB.Close
''''
''''    If glbLinamar Then dblDHours# = 8
''''
''''    'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf.Text))
''''    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf.Text))
''''
''''    intWhereFit& = -1   ' first record can be just less than
''''
''''    For X% = 0 To 16
''''        If medGTServ(X%) > 0 Then
''''            If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
''''                intWhereFit& = X%
''''                If Len(medEntitle(X%)) > 0 Then if_Entitle = True
''''                If Len(medVacation(X%)) > 0 Then if_Vacation = True
''''                Exit For
''''            End If
''''        End If
''''    Next X%
''''
''''    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
''''
''''    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
''''    ' which represents if Sick and Vacation entitlements
''''    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
''''    ' and read on system startup.
''''
''''    ' In this routine we work independantly of SICK/VACATIon entitlement.
''''    '  fglbCompMonthly% - is the independant representation
''''        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
''''        'Procedure modUpdateSelection is used to set
''''        'fglbCompMonthly based on values it finds for global variables
''''        ' and what the user wants to manipulate (sick/Vac)
''''
''''    'optD indicates if Entitlement entered is Daily or yearly based
''''    ' if daily then max entitlement is based on entitlement * hours they work.
''''
''''    ' we have   Entitle = existing entitmenet (stored presently
''''    '           NewEntitle = amount entered onto screen = medentitle(index)
''''    '           EntitleUpd  = value to update record with
''''
''''    If if_Entitle Then
''''        dblNewEntitle# = medEntitle(intWhereFit&)
''''        dblNewMax# = 0
''''        If optD(intWhereFit&) = True Then           ' Entitlements entered in days
''''            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
''''            dblNewEntitle# = dblNewEntitle# * dblDHours#
''''            dblEntitleUpd = dblNewEntitle
''''        End If
''''        If optF(intWhereFit&) = True Then
''''            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# '* dblDHours#
''''            dblNewEntitle# = dblNewEntitle# * dblFTEHours# '* dblDHours#
''''            'FTE HOURS = FTE NUMBER * HOURS PER DAY
''''            dblEntitleUpd = dblNewEntitle
''''        End If
''''        If optH(intWhereFit&) = True Then
''''            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
''''        End If
''''        If fglbCompMonthly Then
''''            dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
''''        Else
''''            dblEntitleUpd = dblNewEntitle
''''        End If
''''
''''         If dblNewMax <> 0 Then          'only do if not zero
''''            If glbCompSerial = "S/N - 2228W" Then  'Simcoe County District Health Unit #3646
''''                                                   'Maximum again Outstanding
''''                If dblEntitleUpd + dblPrevEntitle# - dblEntitleTaken# > dblNewMax Then
''''                    dblEntitleUpd = dblNewMax - dblPrevEntitle# + dblEntitleTaken#
''''                End If
''''            Else
''''                If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
''''                    dblEntitleUpd = dblNewMax - dblPrevEntitle#
''''                End If
''''            End If
''''        End If
''''
''''        DtTm = Now
''''    End If
''''
''''    If if_Vacation Then
''''        VacpcN = medVacation(intWhereFit&)
''''        VacpcO = snapEntitle("ED_VACPC")
''''        VED_DIV = snapEntitle("ED_DIV")
''''        VED_PT = snapEntitle("ED_PT")
''''        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
''''
''''    End If
''''    If if_Entitle Then
''''        snapEntitle(ffieldEntitle$) = dblEntitleUpd       ' base entitlements sic/vacation
''''    End If
''''    snapEntitle.Update
''''
''''    If if_Vacation Then
''''        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
''''        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
''''
''''        SQLQW1 = SQLQW1 & " VALUES('M','N'," & EmpNo& & "," & Val(Format(VacpcN)) & "," & Val(Format(VacpcO))
''''        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
''''        SQLQW1 = SQLQW1 & Date_SQL(Date) & " , '"
''''        SQLQW1 = SQLQW1 & Time$ & "', "
''''        SQLQW1 = SQLQW1 & "'N', "
''''        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
''''        SQLQW1 = SQLQW1 & ")"
''''
''''        gdbAdoIhr001X.Execute SQLQW1
''''    End If
''''
''''lblNextRec:
''''    snapEntitle.MoveNext
''''
''''Wend
''''modUpdateSelection = True
''''MDIMain.panHelp(0).FloodType = 0
''''gdbAdoIhr001.CommitTrans
''''
''''
''''snapEntitle.Close
''''
''''Screen.MousePointer = Default
''''
''''Exit Function
''''
''''modUpdateSelection_Err:
''''If Err = 13 Or Err = 94 Or Err = 3018 Then
''''    Err = 0
''''    Resume Next
''''End If
''''
''''Screen.MousePointer = Default
''''glbFrmCaption$ = Me.Caption
''''glbErrNum& = Err
''''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
''''Screen.MousePointer = Default
''''If gintRollBack% = False Then
''''    'Rollback
''''    Resume Next
''''Else
''''    Unload Me
''''End If
''''End Function
''''
Private Function modZeroSelection()
Dim empNo&
Dim dblEntitle#, dblPrevEntitle#
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblNewMax#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%
Dim xKey, xBankCodeV, xEntV, xFDateV, xTDateV, xBankCodeS, xEntS, xfdateS, xtdateS

' Entitlements are always valued in HOURS - if you enter days then it
'   works out how many hours (based on average Hrswrked/day found in salary master record)
On Error GoTo modZeroSelection_Err

modZeroSelection = False

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    If chkOvtE Then
        Call Overtime_Bank_ZeroOut
    End If
'End If

If chkVacation Or chkSick Then

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
        
        empNo& = snapEntitle("ED_EMPNBR")
        
        DtTm = Now
        xBankCodeS = "": xEntS = 0: xfdateS = "": xtdateS = ""
        If chkSick Then
            xBankCodeV = "SICK"
            xFDateV = snapEntitle("ED_EFDATES")
            xTDateV = snapEntitle("ED_ETDATES")
            If chkZeroCurrent.Value Then
                xComments = "Current Sick Ent. Chg from " & snapEntitle("ED_SICK") & " to 0"
                'Call Append_Accrual(EmpNo&, "SICK", Date, -Val(snapEntitle("ED_SICK") & ""), "Z", xComments)
                Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), -Val(snapEntitle("ED_SICK") & ""), "Z", xComments)
                snapEntitle("ED_SICK") = 0
                snapEntitle("ED_ANNSICK") = 0
                xEntS = 0
            End If
            If chkZeroPrev.Value Then
                xComments = "Prev. Sick Ent. Chg from " & snapEntitle("ED_PSICK") & " to 0"
                'Call Append_Accrual(EmpNo&, "SICK", Date, -Val(snapEntitle("ED_PSICK") & ""), "Z", xComments)
                Call Append_Accrual(empNo&, "SICK", snapEntitle("ED_ETDATES"), -Val(snapEntitle("ED_PSICK") & ""), "Z", xComments)
                snapEntitle("ED_PSICK") = 0
                xEntS = snapEntitle("ED_SICK")
            End If
        End If
        
        xBankCodeV = "": xEntV = 0: xFDateV = "": xTDateV = ""
        
        If chkVacation Then
            xBankCodeV = "VAC"
            xFDateV = snapEntitle("ED_EFDATE")
            xTDateV = snapEntitle("ED_ETDATE")
            If chkZeroCurrent.Value Then
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to 0"
                'Call Append_Accrual(EmpNo&, "VAC", Date, -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
                
                'Ticket #25300: United Way of Lower Mainland - The Rollover has computed new Entitlement Period so
                'now updating Zero Out to Accrual - will compute the Entitlement End Date of Previous period.
                If glbCompSerial = "S/N - 2424W" And Len(cmbAnnMonth) > 0 Then
                    Call Append_Accrual(empNo&, "VAC", DateAdd("d", -1, snapEntitle("ED_EFDATE")), -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
                Else
                    Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
                End If
                snapEntitle("ED_VAC") = 0
                snapEntitle("ED_ANNVAC") = 0
                If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
                    snapEntitle("ED_OTBANK") = 0
                End If
                xEntV = 0
            End If
            If chkZeroPrev.Value Then
                xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to 0"
                'Call Append_Accrual(EmpNo&, "VAC", Date, -Val(snapEntitle("ED_PVAC") & ""), "Z", xComments)
                Call Append_Accrual(empNo&, "VAC", snapEntitle("ED_ETDATE"), -Val(snapEntitle("ED_PVAC") & ""), "Z", xComments)
                snapEntitle("ED_PVAC") = 0
                xEntV = snapEntitle("ED_VAC")
            End If
        End If
    
    
        snapEntitle.Update
        
        If chkVacation Then
            xKey = empNo&
            xKey = xKey & "|" & Format(xFDateV, "dd-mmm-yyyy")
            xKey = xKey & "|" & Format(xTDateV, "dd-mmm-yyyy")
            xKey = xKey & "|VAC"
            xKey = xKey & "|" & xEntV
            xKey = xKey & "|" & Format(xFDateV, "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
            Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker
            DoEvents
        End If
        If chkSick Then
            xKey = empNo&
            xKey = xKey & "|" & Format(xfdateS, "dd-mmm-yyyy")
            xKey = xKey & "|" & Format(xtdateS, "dd-mmm-yyyy")
            xKey = xKey & "|SICK"
            xKey = xKey & "|" & xEntS
            xKey = xKey & "|" & Format(xfdateS, "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
            Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker
            DoEvents
        End If
    
lblNextZRec:
        snapEntitle.MoveNext
        DoEvents
    
    Wend
    modZeroSelection = True
    MDIMain.panHelp(0).FloodType = 0
    snapEntitle.Close
    Screen.MousePointer = DEFAULT
    'gdbAdoIhr001.CommitTrans

End If

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

Private Sub optAccumulate_Click(Value As Integer)
    If optAccumulate Then
        medMaxRollover.Text = ""
        lblMaxRollover.Visible = False
        medMaxRollover.Visible = False
        optDH(0).Visible = False
        optDH(1).Visible = False
    End If
End Sub

Private Sub optD_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optDH_Click(Index As Integer, Value As Integer)
    If Index = 1 Then
        If optDH(1) Then
            MsgBox "Make sure Employee's Hours/Day is specified on the Position screen otherwise the Rollover will skip for that Employee.", vbExclamation, "Maximum Rollover in Days"
        End If
    End If
End Sub

Private Sub optF_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optH_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSick_Click(Value As Integer)
Dim x%

If Value Then
    Call modSetFGlobals("Sick")
End If
Call UpdateEntControls(fglbModifyWhat$ = "Update")

End Sub

Private Sub chkSick_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkVacation_Click(Value As Integer)
If Value Then
    Call modSetFGlobals("Vac")
End If
Call UpdateEntControls(fglbModifyWhat$ = "Update")

End Sub

Private Sub chkVacation_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub UpdateEntControls(ControlsShown) 'js-6Apr99-controls visibility
                                             '  of Update Entitlements controls
End Sub

Private Sub DisplayRule(rsTA As adodb.Recordset)
Dim SQLQ, xOrder, nOrder, aa
Dim rsVE As New adodb.Recordset
Dim x
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
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & clpDept.Text & "' "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "

If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If glbLinamar Then
    If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & clpCode(1).Text & "' "
Else
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(1).Text & "' "
End If
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If clpPT.Text <> "" Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

If Len(cmbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND MONTH(ED_DOH) = " & cmbAnnMonth.ListIndex

End Sub

Private Function AccuValForMulti(empNo, dblEnt) ' Ticket #3304
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

Private Function GetFTEtot(empNo, dblFTE)
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

'''' this is not in use
''''Private Function modUpdateSelectionWHSCC()
'''''Jaddy had some fixes will affect this function. Please talk to her if any change must to be there.
''''Dim EmpNo As Long
''''Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
''''Dim strJob$, dblServiceYears#
''''Dim spt As Variant, varStartDate As Variant, lngRecs&
''''Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
''''Dim dblFTEHours# ', dblFTEHoursTot#
''''Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
''''Dim Msg$, Title$, DgDef As Variant
''''Dim Response%, pct%
''''Dim prec%
''''Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
''''Dim if_Entitle As Boolean, if_Vacation As Boolean
''''Dim ifAnnual As Boolean, dblNewEntAnn#, VacpcNAnn, ifUnionDate As Boolean, ifFirstDate As Boolean, xAsOf 'Frank for WHSCC
''''Dim dblServiceYearsYTD, if_NON As Boolean
''''Dim NoUptSickList As String
''''' Entitlements are always valued in HOURS - if you enter days then it
'''''   works out how many hours (based on average Hrswrked/day found in salary master record)
''''On Error GoTo modUpdateSelectionWHSCC_Err
''''modUpdateSelectionWHSCC = False
''''
''''
''''If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'''''
'''''If fTablHREMP.State <> 0 Then fTablHREMP.Close
'''''fTablHREMP.Open "HREMP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
''''Screen.MousePointer = Default
''''
''''
''''If snapEntitle.BOF And snapEntitle.EOF Then
''''    MsgBox "Employees for this selection do not exist!"
''''    Exit Function
''''Else
''''    lngRecs& = snapEntitle.RecordCount
''''    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
''''    Title$ = "Update Entitlements"
''''    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
''''    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
''''    If Response% = IDNO Then    ' Evaluate response
''''        Exit Function
''''    End If
''''    Screen.MousePointer = HOURGLASS
''''End If
''''MDIMain.panHelp(0).FloodType = 1
''''MDIMain.panHelp(0).FloodPercent = 5
''''
'''''Ticket# 3856
'''''If the employee's Employment Status is one of those on the list,
'''''do not update the employee's sick entitlement for that month. Linda Rowland
''''NoUptSickList = ",BD,CAS,CLIN,CONT,EIS,LTD,MAT,PAR,STUD,"
''''
''''For X% = 0 To 16
''''    If Not IsNumeric(medLTServ(X%)) Then Exit For ' medLTServ(X%) = 0
''''    If Not IsNumeric(medGTServ(X%)) Then
''''      medGTServ(X%) = 0
''''    Else
''''      If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
''''    End If
''''    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
''''Next
''''
''''gdbAdoIhr001.BeginTrans
''''
''''While Not snapEntitle.EOF
''''    prec% = prec% + 1
''''    pct% = Int(100 * (prec% / lngRecs&))
''''    MDIMain.panHelp(0).FloodPercent = pct%
''''    if_Entitle = False
''''    if_Vacation = False
''''
''''    EmpNo& = snapEntitle("ED_EMPNBR")
''''
''''    'Ticket# 3856
''''    If optSickE Then 'For sick
''''        If Not IsNull(snapEntitle("ED_EMP")) Then
''''            If InStr(1, NoUptSickList, "," & Trim(snapEntitle("ED_EMP")) & ",") > 0 Then
''''                GoTo lblNextRec
''''            End If
''''        End If
''''    End If
''''
''''    If IsNull(snapEntitle(ffieldEntitle$)) Then
''''        dblEntitle# = 0
''''    Else
''''        dblEntitle# = snapEntitle(ffieldEntitle$)
''''    End If
''''
''''    If IsNull(snapEntitle(ffieldPEntitle$)) Then
''''        dblPrevEntitle# = 0
''''    Else
''''        dblPrevEntitle# = snapEntitle(ffieldPEntitle$)
''''    End If
''''
''''    If IsNull(snapEntitle(ffieldTEntitle$)) Then
''''        dblTKEEntitle# = 0
''''    Else
''''        dblTKEEntitle# = snapEntitle(ffieldTEntitle$)
''''    End If
''''
''''    spt = snapEntitle("ED_PT")
''''    strDivision$ = snapEntitle("ED_DIV")
''''
''''    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
''''
''''    varStartDate = snapEntitle(fglbWDate$)
''''
''''    Dim rsJOB As New ADODB.Recordset
''''    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
''''    dblDHours# = 0
''''    dblFTEHours# = 0
''''    'This loop is for both of multi positions and not multi
''''    Do Until rsJOB.EOF
''''        If IsNumeric(rsJOB("JH_DHRS")) Then
''''            dblDHours# = dblDHours# + rsJOB("JH_DHRS")
''''            If IsNumeric(rsJOB("JH_FTENUM")) Then
''''                dblFTEHours# = dblFTEHours# + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
''''                'FTE HOURS = FTE NUMBER * HOURS PER DAY
''''            End If
''''        End If
''''        rsJOB.MoveNext
''''    Loop
''''    rsJOB.Close
'''''
'''''    If Not IsNumeric(snapEntitle("JH_DHRS")) Then
'''''        dblDHours# = 0
'''''    Else
'''''        dblDHours# = snapEntitle("JH_DHRS")
'''''    End If
'''''
'''''    If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
'''''        dblFTEHours# = 0
'''''    Else
'''''        dblFTEHours# = snapEntitle("JH_FTENUM")
'''''    End If
'''''    dblFTEHoursTot# = GetFTEtot(EmpNo&, dblFTEHours#) 'For Multi Position, get the Total of FTE for one employee
''''
''''    'Franks Jul 31, 02 for WHSCC
''''    ifAnnual = False
''''    ifUnionDate = False
''''    ifFirstDate = False
''''    If optVacE Then 'Vacation only
''''        'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
''''        dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf.Text))
''''
''''        'dblServiceYearsYTD = (DateDiff("d", varStartDate, CVDate("DEC 31," & Year(dlpAsOf))) / 365) * 12
''''        dblServiceYearsYTD = MonthDiff(CVDate(varStartDate), CVDate("DEC 1," & Year(dlpAsOf)))
''''
''''        If snapEntitle("ED_ORG") = "1866" And snapEntitle("ED_PT") = "FT" Then
''''            If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
''''                ifAnnual = True
''''            End If
''''        End If
''''        If snapEntitle("ED_ORG") = "946" And snapEntitle("ED_PT") = "FT" Then
''''            If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
''''                ifAnnual = True
''''            End If
''''        End If
''''        If snapEntitle("ED_ORG") = "NON" Then ' And snapEntitle("ED_PT") = "FT" Then
''''            If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
''''                'If dblServiceYears# < 120 Then 'Less then 10 years then per month else per year
''''                if_NON = True
''''                If dblServiceYearsYTD < 120 Then 'Less then 10 years months, monthly, otherwise yearly
''''                    ifAnnual = True
''''                End If
''''                'If IsDate(snapEntitle("ED_UNION")) Then
''''                '    ifAnnual = True
''''                '    ifUnionDate = True
''''                'End If
''''                'If IsDate(snapEntitle("ED_FDAY")) Then
''''                '    ifAnnual = True
''''                '    ifFirstDate = True
''''                'End If
''''            End If
''''        End If
''''        If snapEntitle("ED_ORG") = "PHYS" Then 'And snapEntitle("ED_PT") = "FT" Then
''''            If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
''''                'If dblServiceYears# < 120 Then 'Less then 10 years
''''                If dblServiceYearsYTD < 120 Then 'Less then 10 years months, monthly, otherwise yearly
''''                    ifAnnual = True
''''                End If
''''                'If IsDate(snapEntitle("ED_UNION")) Then
''''                '    ifAnnual = True
''''                '    ifUnionDate = True
''''                'End If
''''                'If IsDate(snapEntitle("ED_FDAY")) Then
''''                '    ifAnnual = True
''''                '    ifFirstDate = True
''''                'End If
''''            End If
''''        End If
''''    End If
''''    'Franks Jul 31, 02 for WHSCC
''''
''''    ' dkostka - 08/13/2001 - Changed formula from using number of days / 365 * 12 to using DateDiff
''''    '   directly to get number of months.  We don't get decimals here but the value is always correct.
''''    '   Using the old formula would cause problems sometimes because it assumes all months have an
''''    '   equal number of days, and all years are 365 days.
''''    'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(dlpAsOf)) / 365) * 12
''''    If Not ifAnnual Then
''''        'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
''''        If Not if_NON Then
''''            'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
''''            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf.Text))
''''        Else
''''            dblServiceYears# = dblServiceYearsYTD
''''        End If
''''        intWhereFit& = -1   ' first record can be just less than
''''
''''        For X% = 0 To 16
''''            If medGTServ(X%) > 0 Then
''''                If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
''''                    intWhereFit& = X%
''''                    If Len(medEntitle(X%)) > 0 Then if_Entitle = True
''''                    If Len(medVacation(X%)) > 0 Then if_Vacation = True
''''                    Exit For
''''                End If
''''            End If
''''        Next X%
''''
''''        If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
''''    Else 'Franks Jul 31, 02 for WHSCC
''''        xAsOf = CVDate("Jan 1," & Year(dlpAsOf))
''''        dblNewEntAnn# = 0
''''        VacpcNAnn = 0
''''        intWhereFit& = 0
''''        For z% = 1 To 12
''''            'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
''''            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
''''            'If there is date of Union Date or First Day on Status/Dates screen,
''''            'use the special vacation rules, otherwise use the rules on the Vacation Master screen
''''            If Not (ifUnionDate Or ifFirstDate) Then
''''                For X% = 0 To 16
''''                    If medGTServ(X%) > 0 Then
''''                        If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
''''                            intWhereFit& = X%
''''                            If Len(medEntitle(X%)) > 0 Then
''''                                if_Entitle = True
''''                                dblNewEntAnn# = dblNewEntAnn# + medEntitle(X%)
''''                            End If
''''                            If Len(medVacation(X%)) > 0 Then
''''                                if_Vacation = True
''''                                VacpcNAnn = VacpcNAnn + medVacation(intWhereFit&)
''''                            End If
''''                            Exit For
''''                        End If
''''                    End If
''''                Next X%
''''            Else
''''                If ifUnionDate Then
''''                    If dblServiceYears# >= 0 And dblServiceYears# < 48.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 1.25
''''                    End If
''''                    If dblServiceYears# >= 49 And dblServiceYears# < 239.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 1.67
''''                    End If
''''                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 2.09
''''                    End If
''''                End If
''''                If ifFirstDate Then
''''                    If dblServiceYears# >= 0 And dblServiceYears# < 11.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 1.25
''''                    End If
''''                    If dblServiceYears# >= 12 And dblServiceYears# < 95.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 1.67
''''                    End If
''''                    If dblServiceYears# >= 96 And dblServiceYears# < 239.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 2.09
''''                    End If
''''                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
''''                            if_Entitle = True
''''                            dblNewEntAnn# = dblNewEntAnn# + 2.5
''''                    End If
''''                End If
''''            End If
''''            xAsOf = DateAdd("m", 1, xAsOf)
''''        Next z%
''''    End If 'Franks Jul 31, 02 for WHSCC
''''    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
''''    ' which represents if Sick and Vacation entitlements
''''    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
''''    ' and read on system startup.
''''
''''    ' In this routine we work independantly of SICK/VACATIon entitlement.
''''    '  fglbCompMonthly% - is the independant representation
''''        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
''''        'Procedure modUpdateSelectionWHSCC is used to set
''''        'fglbCompMonthly based on values it finds for global variables
''''        ' and what the user wants to manipulate (sick/Vac)
''''
''''    'optD indicates if Entitlement entered is Daily or yearly based
''''    ' if daily then max entitlement is based on entitlement * hours they work.
''''
''''    ' we have   Entitle = existing entitmenet (stored presently
''''    '           NewEntitle = amount entered onto screen = medentitle(index)
''''    '           EntitleUpd  = value to update record with
''''
''''    If if_Entitle Then
''''        If ifAnnual Then
''''            dblNewEntitle# = dblNewEntAnn#
''''            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
''''                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
''''                dblNewEntitle# = dblNewEntitle# * dblDHours#
''''                dblEntitleUpd = dblNewEntitle
''''            End If
''''            If optF(intWhereFit&) = True Then
''''                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# '* dblFTEHoursTot# * dblDHours#
''''                dblNewEntitle# = dblNewEntitle# * dblFTEHours# '* dblDHours#
''''                dblEntitleUpd = dblNewEntitle
''''            End If
''''            If fglbCompMonthly% Then
''''                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
''''            Else
''''                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
''''                dblEntitleUpd# = dblNewEntitle '+ AccuValForMulti(EmpNo&, dblEntitle#) 'MultiPos Update
''''            End If
''''            If dblNewMax <> 0 Then          'only do if not zero
''''                If optSickE Then 'For sick
''''                    If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
''''                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
''''                    End If
''''                Else
''''                    'If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
''''                    '    dblEntitleUpd = dblNewMax - dblPrevEntitle#
''''                    'End If
''''                    'ticket #3616
''''                    If dblEntitleUpd > dblNewMax Then
''''                        dblEntitleUpd = dblNewMax
''''                    End If
''''                    'ticket #3616
''''                End If
''''            End If
''''        Else
''''            dblNewEntitle# = medEntitle(intWhereFit&)
''''            dblNewMax# = 0
''''            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
''''                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
''''                dblNewEntitle# = dblNewEntitle# * dblDHours#
''''                dblEntitleUpd = dblNewEntitle
''''            End If
''''            If optF(intWhereFit&) = True Then
''''                'If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
''''                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# '* dblFTEHoursTot# * dblDHours#
''''                dblNewEntitle# = dblNewEntitle# * dblFTEHours# '* dblDHours#
''''                dblEntitleUpd = dblNewEntitle
''''            End If
''''            If optH(intWhereFit&) = True Then
''''                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
''''            End If
''''            If fglbCompMonthly% Then
''''                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
''''            Else
''''                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
''''                dblEntitleUpd# = dblNewEntitle '+ AccuValForMulti(EmpNo&, dblEntitle#) 'MultiPos Update
''''            End If
''''
''''            If dblNewMax <> 0 Then          'only do if not zero
''''                If optSickE Then 'For sick
''''                    If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
''''                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
''''                    End If
''''                Else
''''                    'If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
''''                    '    dblEntitleUpd = dblNewMax - dblPrevEntitle#
''''                    'End If
''''                    'ticket #3616
''''                    If dblEntitleUpd > dblNewMax Then
''''                        dblEntitleUpd = dblNewMax
''''                    End If
''''                    'ticket #3616
''''                End If
''''            End If
''''        End If
''''        DtTm = Now
''''    End If
''''
''''    If if_Vacation Then
''''        If Not ifAnnual Then
''''            VacpcN = medVacation(intWhereFit&)
''''        Else   'Franks Jul 31, 02 for WHSCC
''''            VacpcN = VacpcNAnn
''''        End If 'Franks Jul 31, 02 for WHSCC
''''        VacpcO = snapEntitle("ED_VACPC")
''''        VED_DIV = snapEntitle("ED_DIV")
''''        VED_PT = snapEntitle("ED_PT")
''''        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
''''
''''    End If
''''    If if_Entitle Then
''''
''''        If optSickE.Value Then
''''            'For Sick Entitlement update, check the ASL Bank first.
''''            'If ASL Bank is greater than 0, take Repaid ASL from it
''''            'Otherwise, assign the amount to the Sick Entitlement(ED_SICK)
''''            snapEntitle(ffieldEntitle$) = CalcASLRepaid(EmpNo, CVDate(dlpAsOf), dblEntitleUpd, dblNewEntitle, dblEntitle#) 'dblEntitleUpd)
''''        Else
''''            snapEntitle(ffieldEntitle$) = dblEntitleUpd       ' base entitlements sic/vacation
''''        End If
''''    End If
''''    snapEntitle.Update
''''
''''    If if_Vacation Then
''''        ' INSERT INTO HRAUDIT
''''        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
''''        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
''''
''''        ' dkostka - 01/09/01 - Added Val(Format()) around vac pay %, removed quotes.  This prevents the 'data type mismatch' error.
''''        SQLQW1 = SQLQW1 & " VALUES('M','N'," & EmpNo& & "," & Val(Format(VacpcN)) & "," & Val(Format(VacpcO))
''''        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
''''        SQLQW1 = SQLQW1 & IIf(glbSQL, "", "CVDATE") & "('" & Format(Now, "mmm dd,yyyy") & "') , '"
''''        SQLQW1 = SQLQW1 & Time$ & "', "
''''        SQLQW1 = SQLQW1 & "'N', "
''''        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
''''        SQLQW1 = SQLQW1 & ")"
''''
''''        gdbAdoIhr001X.Execute SQLQW1
''''    End If
''''
''''lblNextRec:
''''    snapEntitle.MoveNext
''''
''''Wend
''''modUpdateSelectionWHSCC = True
''''MDIMain.panHelp(0).FloodType = 0
''''gdbAdoIhr001.CommitTrans
''''
'''''fTablHREMP.Close
''''
''''snapEntitle.Close
''''
''''Screen.MousePointer = Default
''''
''''Exit Function
''''
''''modUpdateSelectionWHSCC_Err:
'''''These errors are:
'''''13=type mismatch
'''''94=invalid use of null
'''''3018=couln't find field 'item'
''''If Err = 13 Or Err = 94 Or Err = 3018 Then
''''   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelectionWHSCC" & Chr(10) & "FORM:FUENTITL.FRM"
''''    'commented out by RAUBREY 5/20/97
''''    Err = 0
''''    Resume Next
''''End If
''''
''''Screen.MousePointer = Default
''''glbFrmCaption$ = Me.Caption
''''glbErrNum& = Err
''''Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
''''Screen.MousePointer = Default
''''If gintRollBack% = False Then
''''    'Rollback
''''    Resume Next
''''Else
''''    Unload Me
''''End If
''''End Function
''''
Private Function CalcASLRepaid(xEmpNo, xAsofDate, dblEntUpd, dblNewEnt, dblEnt#) '
Dim rsASL As New adodb.Recordset
Dim rsENT As New adodb.Recordset
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

Private Sub SetSickRules() 'Whscc only
Dim x
    For x = 0 To 16
        medLTServ(x) = ""
        medGTServ(x) = ""
        medEntitle(x) = ""
        optD(x) = False
        optH(x) = False
        optF(x) = True
        medMax(x) = ""
        medVacation(x) = ""
    Next
    medLTServ(0) = 0
    medGTServ(0) = 999
    medEntitle(0) = 1.5
    medMax(0) = 240
End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
'UpdateRight = gSec_Upd_Entitlements
UpdateRight = GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property


Private Sub optReplace_Click(Value As Integer)
    If optReplace Then
        lblMaxRollover.Visible = True
        medMaxRollover.Visible = True
        medMaxRollover.Text = ""
        optDH(0).Visible = True
        optDH(1).Visible = True
    Else
        medMaxRollover.Text = ""
        lblMaxRollover.Visible = False
        medMaxRollover.Visible = False
        optDH(0).Visible = False
        optDH(1).Visible = False
    End If
End Sub
