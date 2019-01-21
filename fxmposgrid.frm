VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMPosGrid 
   Appearance      =   0  'Flat
   Caption         =   "Salary Grid Master"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   960
   ClientWidth     =   10665
   ForeColor       =   &H80000008&
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   10665
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPosDescr2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6270
      MaxLength       =   50
      TabIndex        =   63
      Tag             =   "00-Position Alternate Description"
      Top             =   10110
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   7005
      Left            =   60
      TabIndex        =   38
      Top             =   2850
      Width           =   10125
      Begin VB.TextBox txtNoPos 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         DataField       =   "JB_NBRPOS"
         Height          =   285
         Left            =   1820
         MaxLength       =   3
         TabIndex        =   65
         Tag             =   "10-Number of positions that exist for this job"
         Top             =   1368
         Width           =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpGrid 
         DataField       =   "JB_GRID"
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   300
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         ShowUnassigned  =   1
         TABLName        =   "JBGD"
         Object.Height          =   315
      End
      Begin VB.ComboBox cmbMidPoint 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5820
         TabIndex        =   3
         Tag             =   "01-Mid Point Grid Step number"
         Text            =   "cmbMidPoint"
         Top             =   666
         Width           =   1215
      End
      Begin VB.TextBox txtMidPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "JB_MIDPOINT"
         Height          =   285
         Left            =   6210
         MaxLength       =   2
         TabIndex        =   51
         Top             =   990
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Frame fraGrid 
         Appearance      =   0  'Flat
         Caption         =   "Grid Steps"
         ForeColor       =   &H80000008&
         Height          =   6765
         Left            =   7290
         TabIndex        =   39
         Top             =   120
         Width           =   2130
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S10"
            Height          =   285
            Index           =   10
            Left            =   450
            TabIndex        =   16
            Tag             =   "20-Grid Scales for position"
            Top             =   3111
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S9"
            Height          =   285
            Index           =   9
            Left            =   450
            TabIndex        =   15
            Tag             =   "20-Grid Scales for position"
            Top             =   2792
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S8"
            Height          =   285
            Index           =   8
            Left            =   450
            TabIndex        =   14
            Tag             =   "20-Grid Scales for position"
            Top             =   2473
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S7"
            Height          =   285
            Index           =   7
            Left            =   450
            TabIndex        =   13
            Tag             =   "20-Grid Scales for position"
            Top             =   2154
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S6"
            Height          =   285
            Index           =   6
            Left            =   450
            TabIndex        =   12
            Tag             =   "20-Grid Scales for position"
            Top             =   1835
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S5"
            Height          =   285
            Index           =   5
            Left            =   450
            TabIndex        =   11
            Tag             =   "20-Grid Scales for position"
            Top             =   1516
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S4"
            Height          =   285
            Index           =   4
            Left            =   450
            TabIndex        =   10
            Tag             =   "20-Grid Scales for position"
            Top             =   1197
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S3"
            Height          =   285
            Index           =   3
            Left            =   450
            TabIndex        =   9
            Tag             =   "20-Grid Scales for position"
            Top             =   878
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S2"
            Height          =   285
            Index           =   2
            Left            =   450
            TabIndex        =   8
            Tag             =   "20-Grid Scales for position"
            Top             =   559
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S1"
            Height          =   285
            Index           =   1
            Left            =   450
            TabIndex        =   7
            Tag             =   "21-Grid Scales for position"
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S11"
            Height          =   285
            Index           =   11
            Left            =   450
            TabIndex        =   17
            Tag             =   "20-Grid Scales for position"
            Top             =   3430
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S12"
            Height          =   285
            Index           =   12
            Left            =   450
            TabIndex        =   18
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
            Left            =   450
            TabIndex        =   19
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
            Left            =   450
            TabIndex        =   20
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
            Left            =   450
            TabIndex        =   21
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
            Left            =   450
            TabIndex        =   22
            Tag             =   "20-Grid Scales for position"
            Top             =   5025
            Width           =   1590
            _ExtentX        =   2805
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPayScale 
            DataField       =   "JB_S17"
            Height          =   285
            Index           =   17
            Left            =   450
            TabIndex        =   23
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
            Left            =   450
            TabIndex        =   24
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
            Left            =   450
            TabIndex        =   25
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
            Left            =   450
            TabIndex        =   26
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   80
            Top             =   5070
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "20"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   79
            Top             =   6360
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "19"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   78
            Top             =   6027
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "18"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   77
            Top             =   5708
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "17"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   76
            Top             =   5389
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "12"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   75
            Top             =   3794
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "13"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   74
            Top             =   4113
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "14"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   73
            Top             =   4432
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   72
            Top             =   4751
            Width           =   180
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   50
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   210
            TabIndex        =   49
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   210
            TabIndex        =   48
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   47
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   210
            TabIndex        =   46
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   45
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   44
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   210
            TabIndex        =   43
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   42
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   41
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   40
            Top             =   285
            Width           =   90
         End
      End
      Begin VB.TextBox medFTENum 
         Appearance      =   0  'Flat
         DataField       =   "JB_FTENUM"
         Height          =   285
         Left            =   1820
         TabIndex        =   5
         Tag             =   "10-Number of FTE "
         Top             =   1704
         Width           =   1215
      End
      Begin VB.TextBox medFTEHrs 
         Appearance      =   0  'Flat
         DataField       =   "JB_FTEHRS"
         Height          =   285
         Left            =   1820
         TabIndex        =   6
         Tag             =   "10-FTE Hours/Year"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox comPayPer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "01- Grid Steps - Annual or Hourly"
         Top             =   666
         Width           =   2730
      End
      Begin INFOHR_Controls.CodeLookup clpORG 
         DataField       =   "JB_ORG"
         Height          =   285
         Left            =   1500
         TabIndex        =   4
         Tag             =   "00-Union - Code "
         Top             =   1032
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin VB.Label txtLambtonJob 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1830
         TabIndex        =   71
         Top             =   2940
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Occupation"
         Height          =   195
         Left            =   150
         TabIndex        =   70
         Top             =   2985
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPosFiled 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_NBRFIL"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5220
         TabIndex        =   69
         Top             =   1410
         Width           =   90
      End
      Begin VB.Label lblCountWarn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Warning # Positions < Positions Filled"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3510
         TabIndex        =   68
         Top             =   2460
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label lblPosFilled 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Positions Filled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   67
         Top             =   1413
         Width           =   1035
      End
      Begin VB.Label lblNoPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Positions"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label lblGridC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblSalary 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Grid"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   705
         Width           =   945
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label lblFTENum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label lblFTEHrs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FTE Hours/Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   2085
         Width           =   1395
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "JB_SALCD"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4320
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3465
         TabIndex        =   56
         Top             =   1749
         Width           =   1035
      End
      Begin VB.Label lblTotHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total FTE Hours/Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3465
         TabIndex        =   55
         Top             =   2085
         Width           =   1575
      End
      Begin VB.Label lblFTETotNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_FTETOTNU"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5220
         TabIndex        =   54
         Top             =   1755
         Width           =   90
      End
      Begin VB.Label lblFTETotHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "JB_FTETOTHR"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5220
         TabIndex        =   53
         Top             =   2085
         Width           =   90
      End
      Begin VB.Label lblComRatio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mid-Point"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5100
         TabIndex        =   52
         Top             =   726
         Width           =   840
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4230
      MaxLength       =   25
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5280
      MaxLength       =   25
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   900
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10665
      _Version        =   65536
      _ExtentX        =   18812
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   165
         Width           =   690
      End
      Begin VB.Label lblPosition 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
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
         TabIndex        =   33
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblPosDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
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
         TabIndex        =   32
         Top             =   120
         Width           =   630
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmposgrid.frx":0000
      Height          =   2235
      Left            =   60
      OleObjectBlob   =   "fxmposgrid.frx":0014
      TabIndex        =   0
      Top             =   570
      Width           =   10275
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7500
      Top             =   10920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
      Left            =   8160
      Top             =   10590
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
      PrintFileUseRptDateFmt=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   64
      Top             =   10440
      Width           =   10665
      _Version        =   65536
      _ExtentX        =   18812
      _ExtentY        =   900
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
      Begin VB.CommandButton cmdCountPos 
         Appearance      =   0  'Flat
         Caption         =   "&Count Positions "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   27
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   0
         Width           =   1905
      End
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "JB_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   37
      Top             =   10080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblPositions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POST"
      DataField       =   "JB_CODE"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   36
      Top             =   10290
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   35
      Top             =   10080
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmMPosGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRecords%, fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim LGR_snap As New ADODB.Recordset
Dim snapDiv As New ADODB.Recordset
Dim RDept, RGLNum
Dim rsDATA As New ADODB.Recordset
Dim fglbNew As Boolean
Dim lstMidPoint
'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
'Dim GridStep(11, 2) As Variant
'Dim GridStep(15, 2) As Variant
Dim GridStep(20, 2) As Variant
Dim IfGridStepChange As Boolean
Dim dynSH_Job1 As New ADODB.Recordset

Dim fglbCOMPA#, fglbGRADE$
Dim dblOSalary, dblNewSalary
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, EmpNo&, dblWHours#
Dim oPayP, NPayp, OJOB1, OSalCD

Dim EmpChgErrors As String

Dim SkipResetGridStep As Boolean

Public Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

'Call ST_UPD_MODE(False)  ' reset screen's attributes
Call SET_UP_MODE
'Data1.Recordset.CancelUpdate
'If Not glbSQL Then Call Pause(0.5)
'Data1.Refresh


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRJOBBUD", "Cancel")
Call RollBack   '15June99 js

End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub


Private Sub clpGrid_LostFocus()
If glbLambton Then
    txtLambtonJob = Left(clpGrid, 1) & lblPosition & Mid(clpGrid, 2)
End If
End Sub

Private Sub cmdCountPos_Click()
On Error GoTo CountErr

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If mod_Upd_Pos_Totals(True) Then
        Beep
        MsgBox "Positions Counted"
    End If
    Data1.Refresh
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
Dim a As Integer, Msg As String, INo&

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

If glbVadim Then
    Dim X
    Dim xOccCode
    
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        If Val(medPayScale(X)) > 0 Then
            Call Passing_Salary_Grid_Vadim(X, medPayScale(X), 0, Date, lblPosition, clpGrid)
        End If
    Next
    If glbLambton Then
        xOccCode = Left(clpGrid, 1) & lblPosition & Mid(clpGrid, 2)
    Else
        xOccCode = lblPosition 'getOccCode(xJob)
    End If
    Call Passing_Position_Master_Vadim(xOccCode, "D", lblPosDesc, "")
End If

fglbNew = False
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

'Call ST_UPD_MODE(False)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBBUD", "Delete")
Call RollBack   '15June99 js

End Sub

Public Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo Edit_Err


fglbEditMode% = True


Exit Sub

Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRJOBSKL", "Edit")
Call RollBack   '15June99 js
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
fglbNew = True
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me, rsDATA)
rsDATA.AddNew
clpGrid.Enabled = True
'Data1.Recordset.AddNew
fglbEditMode% = True
lblCNum.Caption = "001"
txtLambtonJob = ""
lblPositions.Caption = glbPos$
lblSalCode = "A"
comPayPer.ListIndex = 0
Call INI_GridStep
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBBUD", "Add")
Call RollBack
End Sub

Public Sub cmdOK_Click()
On Error GoTo OK_Err
Dim I
Dim xGrid
For I = medPayScale.LBound To medPayScale.UBound
    If medPayScale(I).Text = "" Then medPayScale(I).Text = "0.00"
Next I

If Not chkBudgetPos() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

If Not Data1.Recordset.EOF Then
    lstMidPoint = Data1.Recordset("JB_MIDPOINT")    'Hemu
End If
Call Set_Control("U", Me, rsDATA)

rsDATA!JB_MIDPOINT = Val(txtMidPoint)
If rsDATA!JB_MIDPOINT = 0 Then rsDATA!JB_MIDPOINT = 1

xGrid = clpGrid.Text

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh
Data1.Recordset.Find "JB_GRID = '" & xGrid & "'"

EmpChgErrors = ""

If EmpChgErrors <> "" Then
    MsgBox "The following employee salaries could not be changed due to missing 'Hours Per Week' values:" & vbCrLf & EmpChgErrors
End If

SkipResetGridStep = True
DoEvents
Call GridStepChange
DoEvents
SkipResetGridStep = False

fglbNew = False
Call Codes_Master_Integration("POSITION", lblPosition)

Call SET_UP_MODE

Call Display_Value

fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBBUD", "Update")
Call RollBack   '15June99 js
Unload Me
Resume Next
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

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

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0

Me.vbxCrystal.Action = 1

End Sub

Private Sub cmbMidPoint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbMidPoint_LostFocus()
txtMidPoint = cmbMidPoint.ListIndex + 1
End Sub

Private Sub comPayPer_Click()

Select Case comPayPer.ListIndex
    Case 0: lblSalCode.Caption = "A"
    Case 1: lblSalCode.Caption = "H"
    Case Else: lblSalCode.Caption = ""
End Select

End Sub

Private Sub comPayPer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayPer_KeyPress(KeyAscii As Integer)

If comPayPer.Text = "Annual Grid" Then
    lblSalCode.Caption = "A"
ElseIf comPayPer.Text = "Hourly Grid" Then
    lblSalCode.Caption = "H"
End If

End Sub

Private Sub comPayPer_LostFocus()

If comPayPer.ListIndex = 0 Then
    lblSalCode.Caption = "A"
ElseIf comPayPer.ListIndex = 1 Then
    lblSalCode.Caption = "H"
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%

On Error GoTo FLErr
glbOnTop = "FRMMPOSGRID"
Screen.MousePointer = HOURGLASS
If glbPos = "" Then frmJOBS.Show 1
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "Salary Grid Master- " & lblPosition
If glbSyndesis Then
    fraGrid.Caption = "Range"
'    lblGroup.Caption = "Grade"
    vbxTrueGrid.Columns(4).Caption = "Grade"
    vbxTrueGrid.Columns(18).Caption = "Range 1"
    vbxTrueGrid.Columns(19).Caption = "Range 2"
    vbxTrueGrid.Columns(20).Caption = "Range 3"
    vbxTrueGrid.Columns(21).Caption = "Range 4"
    vbxTrueGrid.Columns(22).Caption = "Range 5"
    vbxTrueGrid.Columns(23).Caption = "Range 6 (Mid)"
    vbxTrueGrid.Columns(24).Caption = "Range 7"
    vbxTrueGrid.Columns(25).Caption = "Range 8"
    vbxTrueGrid.Columns(26).Caption = "Range 9"
    vbxTrueGrid.Columns(27).Caption = "Range 10"
    vbxTrueGrid.Columns(28).Caption = "Range 11"
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    vbxTrueGrid.Columns(29).Caption = "Range 12"
    vbxTrueGrid.Columns(30).Caption = "Range 13"
    vbxTrueGrid.Columns(31).Caption = "Range 14"
    vbxTrueGrid.Columns(32).Caption = "Range 15"

    vbxTrueGrid.Columns(33).Caption = "Range 16"
    vbxTrueGrid.Columns(34).Caption = "Range 17"
    vbxTrueGrid.Columns(35).Caption = "Range 18"
    vbxTrueGrid.Columns(36).Caption = "Range 19"
    vbxTrueGrid.Columns(37).Caption = "Range 20"

End If

lblComRatio.Visible = Not glbWFC
cmbMidPoint.Visible = Not glbWFC
comPayPer.Visible = Not glbWFC
lblSalary.Visible = Not glbWFC
fraGrid.Visible = Not glbWFC

Call combMidPoint

comPayPer.AddItem "Annual Grid"
comPayPer.AddItem "Hourly Grid"

Data1.ConnectionString = glbAdoIHRDB

If Not EERetrieve() Then
    Exit Sub        '  modGet it sets fglbRecords
End If

lblGrid(0).Caption = lStr(lblGrid(0))

Call INI_Controls(Me)

Call Display_Value

If glbCompDecHR = 3 Then
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        medPayScale(X).Format = "#,##0.000;(#,##0.000)"
    Next X
End If

If glbCompDecHR = 4 Then
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        medPayScale(X).Format = "#,##0.0000;(#,##0.0000)"
    Next X
End If

If Val(medFTEHrs.Text) = 0 Then
    medFTEHrs = ""
End If

Call SET_UP_MODE

If glbLambton Then
    lblLambtonJob.Visible = True
    txtLambtonJob.Visible = True
End If

Call setCaption(lblUnion)
Call setCaption(lblGridC)
clpGrid.TABLTitle = lStr(lblGridC)

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Budgeted Positions", "Select")
Call RollBack   '15June99 js
Resume Next

End Sub

Public Function EERetrieve() 'StrPos$)
Dim SQLQ$
Dim rsJOB As New ADODB.Recordset
EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr


' out or left join query not updateable - so do straight.
SQLQ$ = "SELECT * FROM HRJOB_GRADE "
SQLQ$ = SQLQ$ & "WHERE JB_CODE = '" & glbPos$ & "' "
SQLQ$ = SQLQ$ & "ORDER BY JB_CODE"

Data1.RecordSource = SQLQ$
Data1.Refresh

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    fglbRecords% = False
Else
    fglbRecords% = True
End If
SQLQ = "SELECT JB_DESCR2,JB_ID FROM HRJOB WHERE JB_CODE='" & glbPos$ & "'"
rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsJOB.EOF Then
    If IsNull(rsJOB("JB_DESCR2")) Then
        txtPosDescr2 = ""
    Else
        txtPosDescr2 = rsJOB("JB_DESCR2")
    End If
End If
rsJOB.Close
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Budgeted Positions", "HRJOBBUD", "SELECT")
Call RollBack   '15June99 js

End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub lblPositions_Change()
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
End Sub


Private Sub lblSalCode_Change()

If Not IsNull(lblSalCode.Caption) Then
    If lblSalCode.Caption = "A" Then
        comPayPer.ListIndex = 0
    ElseIf lblSalCode.Caption = "H" Then
        comPayPer.ListIndex = 1
    End If
End If

End Sub


Private Sub medFTEHrs_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
Private Sub medFTEHrs_Change()
'
'If medFTENum.Text = "" Or medFTENum.Text = "0" Then
'    txtNoPos.Enabled = True 'And cmdOK.Enabled
'Else
'    txtNoPos.Enabled = False
'End If

End Sub



Private Sub medFTEHrs_KeyPress(KeyAscii As Integer)

'If medFTEHrs.Text = "" Then
'    txtNoPos.Enabled = True 'And cmdOK.Enabled
'Else
'    txtNoPos.Enabled = False
'End If

End Sub

Private Sub medFTENum_Change()

'If medFTENum.Text = "" Or medFTENum.Text = "0" Then
'    txtNoPos.Enabled = True 'And cmdOK.Enabled
'Else
'    txtNoPos.Enabled = False
'End If

End Sub

Private Sub medFTENum_GotFocus()

medFTENum.MaxLength = 6 'allows for 6 keystrokes including "."
Call SetPanHelp(ActiveControl)

End Sub

Private Sub medFTENum_KeyPress(KeyAscii As Integer)
'
'If medFTENum.Text = "" Then
'    txtNoPos.Enabled = True 'And cmdOK.Enabled
'Else
'    txtNoPos.Enabled = False
'End If

End Sub



Private Function chkBudgetPos()
Dim SQLQ As String, Msg As String, dd#, PID&, xGrid, xPosCtrl
Dim X
Dim intLastNonZero
chkBudgetPos = False

On Error GoTo chkBudgetPos_Err

If Len(clpGrid) < 1 Then
    MsgBox lStr("Grid Category is a required field")
    clpGrid.SetFocus
    Exit Function
Else
    If clpGrid.Caption = "Unassigned" Then
        MsgBox lStr("Grid Category must be valid")
        clpGrid.SetFocus
        Exit Function
    End If
End If

If Len(clpORG) < 1 Then
    'MsgBox lStr("Union is a required field")
    'clpORG.SetFocus
    'Exit Function
Else
    If clpORG.Caption = "Unassigned" Then
        MsgBox lStr("Union must be valid")
        clpORG.SetFocus
        Exit Function
    End If
End If


If IsNull(rsDATA("JB_ID")) Then
    PID& = 0
Else
    PID& = rsDATA("JB_ID") ' CLng(Val(lblID))
End If
xGrid = clpGrid

If modISDupGrid(glbPos$, xGrid, PID&) Then
    MsgBox lStr("[Grid Category]") & " must be unique"
    clpGrid.SetFocus
    Exit Function
End If
If Len(medFTENum) > 0 Then
    If Not IsNumeric(medFTENum) Then
         MsgBox "You must enter FTE Numeric"
         medFTENum.SetFocus
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

' dkostka - 07/09/2001 - Added check for null Union field to next line, previously if union was null
'   step checks were skipped, as any comparison w/ a NULL is false.
intLastNonZero = 0

If Not glbWFC Or (glbWFC And (clpORG.Text <> "NONE" And clpORG.Text <> "EXEC")) Or (glbWFC And clpORG = "") Then 'Jaddy 10/21/99
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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
                    If intLastNonZero > 0 Then
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
chkBudgetPos = True

Exit Function

chkBudgetPos_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBSKL", "edit/Add")
Call RollBack   '15June99 js

End Function
Private Function modISDupBudgetPosCtrl(Pos$, xPosCtrl, ID&)
Dim SQLQ$
Dim snapBudget As New ADODB.Recordset

modISDupBudgetPosCtrl = True

On Error GoTo modISDupBudget_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HRJOBBUD "
SQLQ$ = SQLQ$ & "Where "
SQLQ$ = SQLQ$ & " (JG_CODE = '" & Pos$ & "' "
SQLQ$ = SQLQ$ & "AND JG_POSCTRLNO = '" & xPosCtrl & "' "
SQLQ$ = SQLQ$ & "AND JG_ID <> " & ID& & ") "
If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapBudget.BOF And snapBudget.EOF Then
    modISDupBudgetPosCtrl = False
End If

Screen.MousePointer = DEFAULT
snapBudget.Close

Exit Function

modISDupBudget_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack   '15June99 js
End Function
Private Function modISDupGrid(Pos$, xGrid, ID&)
Dim SQLQ$
Dim snapGrid As New ADODB.Recordset

modISDupGrid = True

On Error GoTo modISDupGrid_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HRJOB_GRADE "
SQLQ$ = SQLQ$ & " WHERE "
SQLQ$ = SQLQ$ & " (JB_CODE = '" & Pos$ & "' "
SQLQ$ = SQLQ$ & " AND JB_GRID = '" & xGrid & "' "
SQLQ$ = SQLQ$ & "AND JB_ID <> " & ID& & ") "
If snapGrid.State <> 0 Then snapGrid.Close
snapGrid.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapGrid.BOF And snapGrid.EOF Then
    modISDupGrid = False
End If

Screen.MousePointer = DEFAULT
snapGrid.Close

Exit Function

modISDupGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack   '15June99 js

End Function



Private Sub txtNoPos_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT * FROM HRJOB_GRADE "
    SQLQ = SQLQ & "WHERE JB_ID = " & Data1.Recordset!JB_ID

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    lblID = rsDATA!JB_ID
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
Call INI_GridStep
If glbLambton Then
    If Len(clpGrid.Text) > 0 Then
        txtLambtonJob = Left(clpGrid, 1) & lblPosition & Mid(clpGrid, 2)
    End If
End If
End Sub

Private Sub txtMidPoint_Change()
    cmbMidPoint.ListIndex = 0
    If IsNumeric(txtMidPoint) Then cmbMidPoint.ListIndex = Val(txtMidPoint) - 1

End Sub


Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRJOB_GRADE "
        SQLQ = SQLQ & "WHERE JB_CODE = '" & glbPos$ & "' "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X%
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value


If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If Data1.Recordset("JB_SALCD") = "A" Then
        comPayPer.ListIndex = 0
    ElseIf Data1.Recordset("JB_SALCD") = "H" Then
        comPayPer.ListIndex = 1
    Else
        comPayPer.ListIndex = -1
    End If
End If
Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRJOBBUD", "Add")
Call RollBack   '15June99 js

End Sub

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
Dim TF As Boolean, X
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

clpGrid.Enabled = False
txtNoPos.Enabled = TF
cmbMidPoint.Enabled = TF    '
comPayPer.Enabled = TF      '
fraGrid.Enabled = TF        '
medFTEHrs.Enabled = TF      '
medFTENum.Enabled = TF      '
clpORG.Enabled = TF

'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
'For X = 1 To 11
'For X = 1 To 15
For X = 1 To 20
    medPayScale(X).Enabled = TF '
Next

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    glbUserUploadMode = UploadFormWithoutCheck: Unload Me
End If
rr:
End Function


Private Sub INI_GridStep()
Dim X

If SkipResetGridStep Then Exit Sub

If fglbNew Then
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        GridStep(X, 0) = 0
        GridStep(X, 2) = "N"
        
        IfGridStepChange = False
    Next X
    Exit Sub
End If

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X = 1 To 11
    'For X = 1 To 15
    For X = 1 To 20
        GridStep(X, 0) = 0
        GridStep(X, 2) = "N"
        
        IfGridStepChange = False
    Next X
    Exit Sub
End If

'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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
Next X

End Sub

Private Sub combMidPoint()

Dim I, xLev, X

X = cmbMidPoint.ListIndex
cmbMidPoint.Clear

'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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

End Sub

Private Sub GridStepChange()
Dim fTablSalHis As New ADODB.Recordset
Dim lngLastCurrentID&
Dim Msg$, SQLQ
Dim X%, I
Dim xStr
Dim IfChangeMatch As Boolean
Dim Emp_List As New Collection
Dim num
Dim EmpNo&

    On Error GoTo ErrorHandler
    

    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        GridStep(X%, 1) = medPayScale(X%).Text
        If Val(GridStep(X%, 0)) <> Val(GridStep(X%, 1)) Then
            GridStep(X%, 2) = "U"
            IfGridStepChange = True
            If glbVadim Then Call Passing_Salary_Grid_Vadim(X%, GridStep(X%, 0), GridStep(X%, 1), Date, lblPosition, clpGrid.Text)
        End If
    Next X
    
    If fglbNew Then GoTo MarkExit0
    
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
        
        If dynSH_Job1.State <> 0 Then dynSH_Job1.Close
        If fTablSalHis.State <> 0 Then fTablSalHis.Close
        
        fTablSalHis.Open "HR_SALARY_HISTORY", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If glbOracle Then
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY,HRJOB_GRADE WHERE HRJOB_GRADE.JB_CODE = HR_SALARY_HISTORY.SH_JOB AND HRJOB_GRADE.JB_GRID = HR_SALARY_HISTORY.SH_GRID "
            SQLQ = SQLQ & " AND SH_CURRENT <> 0 and SH_JOB = '" & lblPosition & "' AND SH_GRID='" & clpGrid & "'"
        Else
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY INNER JOIN HRJOB_GRADE ON HRJOB_GRADE.JB_CODE = HR_SALARY_HISTORY.SH_JOB AND HRJOB_GRADE.JB_GRID = HR_SALARY_HISTORY.SH_GRID"
            SQLQ = SQLQ & " WHERE SH_CURRENT <> 0 and SH_JOB = '" & lblPosition & "' AND SH_GRID='" & clpGrid & "'"
        End If
        dynSH_Job1.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        If dynSH_Job1.EOF And dynSH_Job1.BOF Then
            GoTo MarkExit
        End If

        Msg$ = "Do you want to update the employee salaries too?"
        X = MsgBox(Msg, 36, "Confirm Update")
        If X <> 6 Then GoTo MarkExit
        
        If CheckGridStepChange = False Then GoTo MarkExit
    End If
    
    glbGridReason = ""
    glbGridEDate = ""
    glbGridNDate = ""
    Load frmForGridStep
    frmForGridStep.Show 1

'MsgBox "glbGridReason = " & glbGridReason
    
    If Len(glbGridReason) = 0 Then GoTo MarkExit
'MsgBox "step 2 "
    Screen.MousePointer = HOURGLASS
    '---- main function
    dynSH_Job1.MoveFirst
    Do While Not dynSH_Job1.EOF
         Emp_List.Add (dynSH_Job1("SH_ID"))
        dynSH_Job1.MoveNext
    Loop
    
'MsgBox "step 3 "
    'Do While Not dynSH_Job1.EOF
    DoEvents
    For num = 1 To Emp_List.count
'MsgBox "step 31 "
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
        
        'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
        'For X% = 1 To 11
        'For X% = 1 To 15
        For X% = 1 To 20
            DoEvents
            If GridStep(X%, 2) = "U" Then ' if update
                If X% = Val(xStr) Then
                    IfChangeMatch = True
                    GoTo Mark1
                End If
            End If
        Next X%
        If Not IfChangeMatch Then GoTo NextRec
Mark1:
        fTablSalHis.MoveFirst
        fTablSalHis.Find "SH_ID = " & lngLastCurrentID&
        OSalary = fTablSalHis("SH_SALARY")
        oPayP = fTablSalHis("SH_PAYP")
        OEDate = fTablSalHis("SH_EDATE")
        ONDate = fTablSalHis("SH_NEXTDAT")
        OJOB1 = fTablSalHis("SH_JOB")
        OSalCD = fTablSalHis("SH_SALCD")
        If Len(fTablSalHis("SH_WHRS")) < 1 Then
            dblWHours# = 0
        Else
            dblWHours# = fTablSalHis("SH_WHRS")
        End If
        fTablSalHis("SH_LTIME") = "T" & I 'Time$
        fTablSalHis("SH_CURRENT") = False
        fTablSalHis.Update
        fTablSalHis.AddNew
        fTablSalHis("SH_COMPNO") = dynSH_Job1("SH_COMPNO")
        fTablSalHis("SH_EMPNBR") = dynSH_Job1("SH_EMPNBR")
        fTablSalHis("SH_EDATE") = CVDate(glbGridEDate) '(txtEDate)
        fTablSalHis("SH_CURRENT") = True
        fTablSalHis("SH_SDATE") = dynSH_Job1("SH_SDATE")
        fTablSalHis("SH_SALCD") = dynSH_Job1("SH_SALCD")
        fTablSalHis("SH_PAYROLL_ID") = dynSH_Job1("SH_PAYROLL_ID")
        fTablSalHis("SH_WHRS") = dynSH_Job1("SH_WHRS")
        fTablSalHis("SH_PAYP") = dynSH_Job1("SH_PAYP")
        fTablSalHis("SH_PAYP_TABLE") = dynSH_Job1("SH_PAYP_TABLE")
        fTablSalHis("SH_SREAS_TABLE") = dynSH_Job1("SH_SREAS_TABLE")
        
        dblOSalary = dynSH_Job1("SH_SALARY")
    
        dblNewSalary = Val(GridStep(X%, 1))
        Call GetNewStepSalary(dblNewSalary, X%)
        ' dkostka - 01/29/2002 - Added list of employees that couldn't be changed for grid step changes.
        If dblNewSalary = -1 Then
            ' Couldn't change salary (no WHRS).  Abort and add to the list of errors.
            EmpChgErrors = EmpChgErrors & dynSH_Job1("SH_EMPNBR") & vbCrLf
            fTablSalHis.CancelUpdate
            GoTo NextRec
        End If
        fTablSalHis("SH_SALARY") = Round2DEC(dblNewSalary)
        If IsDate(glbGridNDate) Then
            fTablSalHis("SH_NEXTDAT") = glbGridNDate
        Else
            If IsDate(ONDate) Then
                If glbLambton Then
                    If CVDate(ONDate) >= CVDate(glbGridEDate) Then
                        fTablSalHis("SH_NEXTDAT") = ONDate
                    End If
                ElseIf CVDate(ONDate) > CVDate(glbGridEDate) Then
                    fTablSalHis("SH_NEXTDAT") = ONDate
                End If
            End If
        End If
        fTablSalHis("SH_JOB") = dynSH_Job1("SH_JOB")
        fTablSalHis("SH_GRID") = dynSH_Job1("SH_GRID")
        fTablSalHis("SH_JOB_ID") = dynSH_Job1("SH_JOB_ID")

        Call modSetCOMPA_GRADE(dblNewSalary) ' sets fglbCOMPA#, and fglbGRADE
        fTablSalHis("SH_COMPA") = Round(fglbCOMPA#, 2)
        fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")

        fTablSalHis("SH_SREAS1") = glbGridReason ' clpCode(4).Text
        If dblOSalary <> 0 Then fTablSalHis("SH_SALPC1") = (dblNewSalary - dblOSalary) / dblOSalary
        fTablSalHis("SH_SALCHG1") = dblNewSalary - dblOSalary

        fTablSalHis("SH_LDATE") = Now
        fTablSalHis("SH_LTIME") = Time$ '"T" & I
        fTablSalHis("SH_LUSER") = glbUserID
        fTablSalHis.Update
        Call Transfer_Salary(fTablSalHis)
        
        Call updFollow("U") 'If Next Review Date enter 0
        
        Call updBenefitForSalDEPN(EmpNo&)
        Call Employee_Master_Integration(EmpNo&)
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

Private Sub GetNewStepSalary(dblNewSalary, X%)
'Days added by Bryan 30/Sep/05 Ticket#9354
Dim dblHoursPerWeek#

dblHoursPerWeek# = dynSH_Job1("SH_WHRS")

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
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If dblHoursPerWeek# = 0 Then
            dblNewSalary = -1
        Else
            If GetLeapYear(Year(Date)) Then
                dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * 366 / (dblHoursPerWeek# * 52)
            Else
                dblNewSalary = Data1.Recordset("JB_S" & Format(X%, "##")) * 365 / (dblHoursPerWeek# * 52)
            End If
        End If
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
    End If
End If
End Sub

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

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

'Days added by Bryan 30/Sep/05 Ticket#9354
ssalary@ = dblNewSalary

dblHoursPerWeek# = dynSH_Job1("SH_WHRS")

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
            dblSsalary# = -1
        Else
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = dblNewSalary * 366 / (dblHoursPerWeek# * 52)
            Else
                dblSsalary# = dblNewSalary * 365 / (dblHoursPerWeek# * 52)
            End If
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
    End If
End If

 ' set COMPA RATIO
 'laura 03/23/98

'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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
SQLQ = "SELECT * FROM HRJOB_GRADE"
SQLQ = SQLQ & " WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
SQLQ = SQLQ & " AND JB_GRID='" & dynSH_Job1("SH_GRID") & "'"
snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic

fglbGRADE$ = "00"
xSalGrade = dblNewSalary

'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 366 / (dblHoursPerWeek# * 52)
                Else
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 365 / (dblHoursPerWeek# * 52)
                End If
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
Dim xADD, xPT, xDiv
Dim TB As New ADODB.Recordset
On Error GoTo AUDIT_ERR
AUDITSALY = False


TB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
TB.MoveFirst
TB.Find "ED_EMPNBR = " & EmpNo&
If Not TB.EOF Then
    xPT = TB("ED_PT")
    xDiv = TB("ED_DIV")
Else
    xPT = ""
    xDiv = ""
End If
Dim strFields As String
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_JOB, AU_SALCD, AU_WHRS, "
strFields = strFields & "AU_SEDATE , AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

If OSalary <> NSalary Then GoTo MODUPD
If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD
GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR": TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL": TA("AU_EARN_TABL") = "EARN"
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

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = EmpNo&
TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"
'If glbSoroc Or glbSyndesis Then
    Dim rsEMP As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & EmpNo&
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then TA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
    End If
    rsEMP.Close
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
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HR_SALARY_HISTORY.SH_GRID "
        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0"
    Else
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY INNER JOIN HR_SALARY_HISTORY "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HR_SALARY_HISTORY.SH_GRID "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
    End If
    SQLQ = SQLQ & " AND JH_JOB='" & lblPosition & "' AND JH_GRID= '" & clpGrid & "'"
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
    SQLQ = "SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE (SH_GRADE='0' OR SH_GRADE='00' OR SH_GRADE IS NULL) "
    SQLQ = SQLQ & " AND SH_JOB='" & lblPosition & "' AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND SH_GRID= '" & clpGrid & "'"
    
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do Until rsMain.EOF
        NotUpdate = NotUpdate & ShowEmpnbr(rsMain("SH_EMPNBR")) & " - Previous salary was at Step 00" & vbCrLf
        rsMain.MoveNext
    Loop
    rsMain.Close
    
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_JOB='" & lblPosition & "'"
    SQLQ = SQLQ & " AND JH_GRID= '" & clpGrid & "'"
    SQLQ = SQLQ & " AND JH_CURRENT<>0"
    
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Do Until rsMain.EOF
        If InStr(NotUpdate, ShowEmpnbr(rsMain("JH_EMPNBR")) & " ") = 0 Then WillUpdate = WillUpdate & ShowEmpnbr(rsMain("JH_EMPNBR")) & vbCrLf
        rsMain.MoveNext
    Loop
    rsMain.Close
    
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
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY,HRJOB_GRADE WHERE HRJOB_GRADE.JB_CODE = HR_SALARY_HISTORY.SH_JOB  AND HRJOB_GRADE.JB_GRID = HR_SALARY_HISTORY.SH_GRID"
        SQLQ = SQLQ & " AND SH_CURRENT <> 0 and SH_JOB = '" & lblPosition & "'"
        SQLQ = SQLQ & " AND SH_GRID= '" & clpGrid & "'"
    Else
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY INNER JOIN HRJOB_GRADE ON HRJOB_GRADE.JB_CODE = HR_SALARY_HISTORY.SH_JOB AND HRJOB_GRADE.JB_GRID = HR_SALARY_HISTORY.SH_GRID"
        SQLQ = SQLQ & " WHERE SH_CURRENT <> 0 and SH_JOB = '" & lblPosition & "'"
         SQLQ = SQLQ & " AND SH_GRID= '" & clpGrid & "'"
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
            dblSalary = fTablSalHis("SH_SALARY")

            'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
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
Dim xGridCode
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
SQLQ = "UPDATE HRJOB_GRADE "
SQLQ = SQLQ & " SET "
SQLQ = SQLQ & " JB_NBRFIL = 0, "
SQLQ = SQLQ & " JB_FTETOTNU = 0, "
SQLQ = SQLQ & " JB_FTETOTHR = 0 "

gdbAdoIhr001.Execute SQLQ
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If

'
SQLQ = "SELECT JH_JOB, JH_GRID,COUNT(JH_EMPNBR) AS NoPosFilled FROM HR_JOB_HISTORY WHERE (JH_CURRENT <> 0) GROUP BY JH_JOB,JH_GRID"

snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 5
If updPCtComp Then
    MDIMain.panHelp(0).FloodPercent = pct#
End If


'rsHrJob.Open "HRJOB", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
''rsHrJob.Index = "PrimaryKey"
If Not snapJobCount.BOF And Not snapJobCount.EOF Then
    snapJobCount.MoveLast
    rcount& = snapJobCount.RecordCount
    snapJobCount.MoveFirst
    ipct# = 20 / rcount&
End If
'
'
'
While Not snapJobCount.BOF And Not snapJobCount.EOF
    Job$ = snapJobCount("JH_JOB")
    xGridCode = snapJobCount("JH_GRID")
    JobCount& = snapJobCount("NoPosFilled")
    
    SQLQ = "SELECT JB_NBRFIL FROM HRJOB_GRADE  "
    SQLQ = SQLQ & " WHERE JB_CODE = '" & Job$ & "'"
    SQLQ = SQLQ & " AND JB_GRID = '" & xGridCode & "'"
    rsHRJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRJob.EOF Then
        rsHRJob("JB_NBRFIL") = JobCount&
        rsHRJob.Update
    End If
    rsHRJob.Close
    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#

    End If

    snapJobCount.MoveNext
Wend

snapJobCount.Close

SQLQ = "SELECT JH_JOB, JH_GRID, Sum(JH_FTENUM) AS FTENumTot"
SQLQ = SQLQ & " From HR_JOB_HISTORY"
SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
SQLQ = SQLQ & " GROUP BY JH_JOB,JH_GRID "


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
    xGridCode = snapFTENum("JH_Grid")
    If IsNull(snapFTENum("FTENumTot")) Then    'laura 03/05/98
      FTENum# = 0
    Else
      FTENum# = snapFTENum("FTENumTot")
    End If
    
    SQLQ = "SELECT JB_FTETOTNU FROM HRJOB_GRADE  "
    SQLQ = SQLQ & " WHERE JB_CODE = '" & Job$ & "'"
    SQLQ = SQLQ & " AND JB_GRID = '" & xGridCode & "'"
    rsHRJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsHRJob.EOF Then
        rsHRJob("JB_FTETotNu") = FTENum#
        rsHRJob.Update
    End If
    rsHRJob.Close
    pct# = pct# + ipct#
    spct% = CInt(pct#)

    If updPCtComp Then
        MDIMain.panHelp(0).FloodPercent = pct#
    End If

    snapFTENum.MoveNext
Wend

snapFTENum.Close

SQLQ = "qry_Count_FTEHrs"

SQLQ = "SELECT JH_JOB, JH_GRID, Sum(JH_FTEHRS) AS FTEHrsTot"
SQLQ = SQLQ & " From HR_JOB_HISTORY"
SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
SQLQ = SQLQ & " GROUP BY JH_JOB,JH_GRID "

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
    xGridCode = snapFTEHrs("JH_Grid")
    If IsNull(snapFTEHrs("FTEHrsTot")) Then     'laura 03/04/98
      FTEHrs& = 0
    Else
      FTEHrs& = snapFTEHrs("FTEHrsTot")
    End If
    
    SQLQ = "SELECT JB_FTETOTHR FROM HRJOB_GRADE  "
    SQLQ = SQLQ & " WHERE JB_CODE = '" & Job$ & "'"
    SQLQ = SQLQ & " AND JB_GRID = '" & xGridCode & "'"
    rsHRJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsHRJob.EOF Then
        rsHRJob("JB_FTETOTHR") = FTEHrs&
        rsHRJob.Update
    End If
    rsHRJob.Close
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
   
    If isChanged_Salary(HRSalary, OSalary, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    If glbVadim And UpdateAudit Then
        Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWHrs, xEmpnbr, xPayrollID, , xNiagaraWHRS)
    End If
    If isChanged_Field(HRChanges, OEDate, rsNew("SH_EDATE")) Then UpdateAudit = True
    If isChanged_Field(HRChanges, ONDate, rsNew("SH_NEXTDAT")) Then UpdateAudit = True
    Call Passing_Changes(HRChanges, Salary, "M", Date, xEmpnbr, xPayrollID)

End Sub
