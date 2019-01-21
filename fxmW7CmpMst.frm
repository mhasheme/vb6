VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMW7CmpMst 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Form 7 Employer Information"
   ClientHeight    =   8865
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   11280
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
   ScaleHeight     =   8865
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcEmplrInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8175
      ScaleWidth      =   10575
      TabIndex        =   26
      Top             =   240
      Width           =   10575
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         DataField       =   "EY_CITY"
         DataSource      =   " "
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "00-City/Town"
         Top             =   3889
         Width           =   3420
      End
      Begin VB.TextBox txtMailAddress 
         Appearance      =   0  'Flat
         DataField       =   "EY_MAIL_ADDRESS"
         DataSource      =   " "
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
         Left            =   2400
         MaxLength       =   45
         TabIndex        =   8
         Tag             =   "00-Mailing Address"
         Top             =   3513
         Width           =   4635
      End
      Begin VB.TextBox txtFirmAcctNo 
         Appearance      =   0  'Flat
         DataField       =   "EY_FIRM_ACCT_NUM"
         Height          =   285
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "00-Firm or Account #"
         Top             =   2385
         Width           =   1680
      End
      Begin VB.TextBox txtDescBusiness 
         Appearance      =   0  'Flat
         DataField       =   "EY_BUSINESS_DESC"
         DataSource      =   " "
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
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   15
         Tag             =   "00-Description of Business Activity"
         Top             =   5400
         Width           =   4635
      End
      Begin VB.TextBox txtTradeLegalName 
         Appearance      =   0  'Flat
         DataField       =   "EY_TRADLEGAL_NAME"
         DataSource      =   " "
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
         Left            =   2400
         MaxLength       =   45
         TabIndex        =   7
         Tag             =   "00-Trade and Legal Name (if different provide both)"
         Top             =   3137
         Width           =   4635
      End
      Begin VB.TextBox txtRateGroup 
         Appearance      =   0  'Flat
         DataField       =   "EY_RATE_GRP_NUM"
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "00-Rate Group Number"
         Top             =   2761
         Width           =   1440
      End
      Begin VB.TextBox txtClassUnit 
         Appearance      =   0  'Flat
         DataField       =   "EY_CLASS_UNIT_CODE"
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
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "00-Classification Unit Code"
         Top             =   2761
         Width           =   1455
      End
      Begin VB.Frame frFirmWorkers 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8160
         TabIndex        =   34
         Top             =   2400
         Width           =   1695
         Begin Threed.SSOption optWorkerYesNo 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   " Yes"
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
         Begin Threed.SSOption optWorkerYesNo 
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   0
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   " No"
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
      Begin VB.TextBox txtFirmAcct 
         Appearance      =   0  'Flat
         DataField       =   "EY_FIRM_ACCT"
         Height          =   285
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "00-Firm or Account"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtFirmWorkers 
         Appearance      =   0  'Flat
         DataField       =   "EY_WKER_GRT_20"
         Height          =   285
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "00-Firm Workers > 20?"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Frame Frame1 
         Caption         =   "Branch Address"
         Height          =   1815
         Left            =   0
         TabIndex        =   27
         Top             =   6000
         Width           =   10455
         Begin VB.TextBox txtBranchCity 
            Appearance      =   0  'Flat
            DataField       =   "EY_WKER_BRNC_CITY"
            DataSource      =   " "
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
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "00-City/Town"
            Top             =   1080
            Width           =   3420
         End
         Begin VB.TextBox txtBranchAdd 
            Appearance      =   0  'Flat
            DataField       =   "EY_WKER_BRNC_ADDR"
            DataSource      =   " "
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
            Left            =   1800
            MaxLength       =   45
            TabIndex        =   17
            Tag             =   "00-Branch Address where worker is based (if different from mailing address - no abbreviations)"
            Top             =   720
            Width           =   4635
         End
         Begin VB.CheckBox chkSameAsMailing 
            Caption         =   "Same as Mailing Address"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Tag             =   "Completed"
            Top             =   360
            Width           =   2715
         End
         Begin INFOHR_Controls.CodeLookup clpBranchProv 
            DataField       =   "EY_WKER_BRNC_PROV"
            Height          =   285
            Left            =   1485
            TabIndex        =   19
            Tag             =   "31-Province - Code"
            Top             =   1440
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   4
         End
         Begin MSMask.MaskEdBox medBranchPCode 
            DataField       =   "EY_WKER_BRNC_PCODE"
            Height          =   285
            Left            =   6360
            TabIndex        =   20
            Tag             =   "00-Postal Code"
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   10
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
         Begin VB.Label lblTitle 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Province"
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
            TabIndex        =   31
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Code"
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
            Left            =   4995
            TabIndex        =   30
            Top             =   1485
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Address"
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
            TabIndex        =   29
            Top             =   765
            Width           =   1680
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "City / Town"
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
            TabIndex        =   28
            Top             =   1125
            Width           =   1320
         End
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fxmW7CmpMst.frx":0000
         Height          =   2175
         Left            =   0
         OleObjectBlob   =   "fxmW7CmpMst.frx":0014
         TabIndex        =   35
         Top             =   0
         Width           =   10455
      End
      Begin Threed.SSOption optFirmAcct 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   2400
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Account #"
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
      Begin INFOHR_Controls.CodeLookup clpProv 
         DataField       =   "EY_PROV"
         Height          =   285
         Left            =   2085
         TabIndex        =   10
         Tag             =   "31-Province - Code"
         Top             =   4265
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin MSMask.MaskEdBox medPCode 
         DataField       =   "EY_PCODE"
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Tag             =   "00-Postal Code"
         Top             =   4265
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   10
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
      Begin MSMask.MaskEdBox medTelephone 
         DataField       =   "EY_PHONE"
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Tag             =   "11-Telephone Number"
         Top             =   4641
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medFax 
         DataField       =   "EY_FAX"
         Height          =   285
         Left            =   6960
         TabIndex        =   13
         Tag             =   "10-Fax Number"
         Top             =   4641
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medAltTelephone 
         DataField       =   "EY_WKER_BRNC_PHONE"
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Tag             =   "11-Alternate Telephone Number"
         Top             =   5017
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optFirmAcct 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   2400
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Firm "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   47
         Top             =   3560
         Width           =   1680
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Trade and Legal Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   3183
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "City / Town"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   3937
         Width           =   1320
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Does your firm have 20 or more workers?"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   4440
         TabIndex        =   44
         Top             =   2430
         Width           =   3600
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   43
         Top             =   4310
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   42
         Top             =   4314
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Left            =   5520
         TabIndex        =   41
         Top             =   4686
         Width           =   255
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   4691
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Business Activity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   39
         Top             =   5445
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate Telephone"
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
         Left            =   120
         TabIndex        =   38
         Top             =   5068
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Group Number"
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
         Left            =   120
         TabIndex        =   37
         Top             =   2806
         Width           =   1425
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification Unit"
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
         Left            =   5520
         TabIndex        =   36
         Top             =   2806
         Width           =   1260
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   7935
      LargeChange     =   315
      Left            =   10920
      Max             =   100
      SmallChange     =   315
      TabIndex        =   25
      Top             =   0
      Width           =   340
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5760
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EY_LUSER"
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8640
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EY_LDATE"
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
      Index           =   0
      Left            =   8280
      MaxLength       =   12
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8625
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EY_LTIME"
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
      Left            =   9030
      MaxLength       =   8
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8640
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7800
      Top             =   8520
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
      GridSource      =   "vbxTrueGrid"
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "EY_COMPNO"
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
      Left            =   4920
      TabIndex        =   23
      Top             =   8640
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmMW7CmpMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum

Private Function chkMW7CmpMaster()
Dim Msg As String
Dim x%, xchk

chkMW7CmpMaster = False

If Len(txtFirmAcctNo.Text) < 1 Then
    MsgBox "Firm/Account # is a required field"
    txtFirmAcctNo.SetFocus
    Exit Function
End If

If Len(txtTradeLegalName.Text) < 1 Then
    MsgBox "Trade/Legal Name is a required field"
    txtTradeLegalName.SetFocus
    Exit Function
End If

If Len(txtMailAddress.Text) < 1 Then
    MsgBox "Mailing Address is a required field"
    txtMailAddress.SetFocus
    Exit Function
End If

If Len(txtCity.Text) < 1 Then
    MsgBox "City / Town is a required field"
    txtCity.SetFocus
    Exit Function
End If

If Len(clpProv.Text) < 1 Then
    MsgBox "Province is a required field"
    clpProv.SetFocus
    Exit Function
Else
    If clpProv.Caption = "Unassigned" Then
        MsgBox "Invalid Province"
        clpProv.SetFocus
        Exit Function
    End If
End If

If Len(medPCode) < 1 Then
    MsgBox "Postal Code is a required field"
    medPCode.SetFocus
    Exit Function
End If

If Len(medTelephone) < 1 Then
    MsgBox "Telephone Number is a required field"
    medTelephone.SetFocus
    Exit Function
End If

If Len(txtDescBusiness.Text) < 1 Then
    MsgBox "Description of Business is a required field"
    txtDescBusiness.SetFocus
    Exit Function
End If

'Jerry does not want these fields to be mandatory
'If Len(txtBranchAdd.Text) < 1 Then
'    MsgBox "Branch Address is a required field"
'    txtBranchAdd.SetFocus
'    Exit Function
'End If
'
'If Len(txtBranchCity.Text) < 1 Then
'    MsgBox "Branch City / Town is a required field"
'    txtBranchCity.SetFocus
'    Exit Function
'End If
'
'If Len(clpBranchProv.Text) < 1 Then
'    MsgBox "Branch Province is a required field"
'    clpBranchProv.SetFocus
'    Exit Function
'Else
    If clpBranchProv.Caption = "Unassigned" Then
        MsgBox "Invalid Branch Province"
        clpBranchProv.SetFocus
        Exit Function
    End If
''End If
'
'If Len(medBranchPCode) < 1 Then
'    MsgBox "Postal Code is a required field"
'    medBranchPCode.SetFocus
'    Exit Function
'End If


xchk = False

chkMW7CmpMaster = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate

Call Display_Value

chkSameAsMailing.Value = False  'reset

'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_OHS_COMPANY_MASTER", "Cancel")
Call RollBack '09June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'Data1.Recordset.Delete
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_COMPANY_MASTER", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

txtFirmAcctNo.SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OHS_COMPANY_MASTER", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

lblCNum.Caption = "001"

fglbNew = True

Call SET_UP_MODE

chkSameAsMailing.Value = False
optFirmAcct(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_COMPANY_MASTER", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%

On Error GoTo cmdOK_Err

If Not chkMW7CmpMaster() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)

If optFirmAcct(0).Value = True Then txtFirmAcct.Text = "A"
If optFirmAcct(1).Value = True Then txtFirmAcct.Text = "F"

If optWorkerYesNo(0).Value = True Then txtFirmWorkers.Text = "1"
If optWorkerYesNo(1).Value = True Then txtFirmWorkers.Text = "0"

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

fglbNew = False

Call SET_UP_MODE

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_COMPANY_MASTER", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Employer Information"
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

RHeading = "Employer Information"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkSameAsMailing_Click()
    If chkSameAsMailing Then
        txtBranchAdd.Text = txtMailAddress.Text
        txtBranchCity.Text = txtCity.Text
        clpBranchProv.Text = clpProv
        medBranchPCode.Text = medPCode.Text
    End If
End Sub

Private Sub clpBranchProv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpProv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_COMPANY_MASTER", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%

Me.Show
glbOnTop = "FRMMW7CMPMST"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_OHS_COMPANY_MASTER ORDER BY EY_FIRM_ACCT_NUM,EY_ID"
Data1.Refresh

Call setRptCaption(Me)

Screen.MousePointer = DEFAULT

'Call Display_Value

Call ST_UPD_MODE(False)

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

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

Private Sub Form_Resize()
Dim c As Long

On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 9000 Then
        scrControl.Value = 0
        
        pcEmplrInfo.Top = 120
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 1000
        
        scrControl.Max = 5000
    End If

'    'Horizontal Scroll
'    scrHScroll.Width = Me.Width - 200
'    If Me.Width >= 11190 Then '9700 Then
'        scrHScroll.Value = 0
'        scrHScroll.Visible = False
'    Else
'        scrHScroll.Visible = True
'        scrHScroll.Top = Me.Height - 700
'        scrHScroll.Width = Me.Width - 120
'    End If
    
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Form 7 Employer Info.", "Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
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

fUPMode = TF
'vbxTrueGrid.Enabled = FT

optFirmAcct(0).Enabled = TF
optFirmAcct(1).Enabled = TF
txtFirmAcctNo.Enabled = TF
txtRateGroup.Enabled = TF
txtClassUnit.Enabled = TF
txtTradeLegalName.Enabled = TF
txtMailAddress.Enabled = TF
txtCity.Enabled = TF
clpProv.Enabled = TF
medPCode.Enabled = TF
medTelephone.Enabled = TF
medFax.Enabled = TF
txtDescBusiness.Enabled = TF
optWorkerYesNo(0).Enabled = TF
optWorkerYesNo(1).Enabled = TF
txtBranchAdd.Enabled = TF
txtBranchCity.Enabled = TF
clpBranchProv.Enabled = TF
medBranchPCode.Enabled = TF
medAltTelephone.Enabled = TF

End Sub

Private Sub medAltTelephone_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medBranchPCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medBranchPCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub medFax_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub medTelephone_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optFirmAcct_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optWorkerYesNo_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
pcEmplrInfo.Top = 120 - scrControl.Value
End Sub

Private Sub txtBranchAdd_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtBranchCity_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCity_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtClassUnit_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDescBusiness_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFirmAcct_Change()
If txtFirmAcct.Text = "A" Then optFirmAcct(0).Value = True
If txtFirmAcct.Text = "F" Then optFirmAcct(1).Value = True
End Sub

Private Sub txtFirmAcctNo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFirmWorkers_Change()
If txtFirmWorkers.Text = "1" Or txtFirmWorkers.Text = "-1" Then optWorkerYesNo(0).Value = True
If txtFirmWorkers.Text = "0" Then optWorkerYesNo(1).Value = True
End Sub

Private Sub txtMailAddress_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRateGroup_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTradeLegalName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_COMPANY_MASTER "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If

Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OHS_COMPANY_MASTER", "Select")
Call RollBack '09June99 js

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
   
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_COMPANY_MASTER where EY_ID= " & Data1.Recordset!EY_ID
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_HSW7CmpMst And glbWSIBModule
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
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub
