VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEBANK 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Banking Information"
   ClientHeight    =   8490
   ClientLeft      =   135
   ClientTop       =   1365
   ClientWidth     =   12840
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel fraUSA 
      Height          =   3405
      Left            =   11640
      TabIndex        =   131
      Top             =   3480
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   6006
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.TextBox txtStatusFalg3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ED_PENSION"
         Height          =   285
         Left            =   2340
         MaxLength       =   1
         TabIndex        =   141
         Tag             =   "00-Status Flag 3"
         Top             =   1800
         Width           =   660
      End
      Begin VB.TextBox txtFedExemp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ED_UIC"
         Height          =   285
         Left            =   2340
         MaxLength       =   2
         TabIndex        =   137
         Tag             =   "00-Federal Tax Exemptions"
         Top             =   705
         Width           =   900
      End
      Begin VB.TextBox txtStateExemption 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ED_CPP"
         Height          =   285
         Left            =   7560
         MaxLength       =   1
         TabIndex        =   136
         Tag             =   "01-State Tax Exemptions"
         Top             =   705
         Width           =   900
      End
      Begin VB.TextBox txtFedMarry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "ED_GROSSCD"
         Height          =   285
         Left            =   4080
         MaxLength       =   1
         TabIndex        =   135
         Tag             =   "00-Income Tax Applicable- Y/N"
         Top             =   360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtStateMarry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "ED_WCBCODE"
         Height          =   285
         Left            =   8760
         MaxLength       =   1
         TabIndex        =   134
         Tag             =   "00-State Marital Status"
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ComboBox cboFedMarry 
         Height          =   315
         Left            =   2340
         TabIndex        =   133
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboStateMarry 
         Height          =   315
         Left            =   7560
         TabIndex        =   132
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox medFedExtra 
         DataField       =   "ED_ExtraTax"
         Height          =   285
         Left            =   2340
         TabIndex        =   138
         Tag             =   "20-Extra Tax on Federal Form"
         Top             =   1095
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
      Begin MSMask.MaskEdBox medFedExtraPC 
         DataField       =   "ED_ExtraTaxPC"
         Height          =   285
         Left            =   2340
         TabIndex        =   139
         Tag             =   "10-Extra Tax Percentage on Federal Form"
         Top             =   1455
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medStateExtraPC 
         DataField       =   "ED_TD1DOL"
         Height          =   285
         Left            =   7560
         TabIndex        =   140
         Tag             =   "20-Amount as found on State"
         Top             =   1455
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
      Begin MSMask.MaskEdBox medStateExtra 
         DataField       =   "ED_TD3"
         Height          =   285
         Left            =   7560
         TabIndex        =   147
         Tag             =   "20-Extra Tax on State Form"
         Top             =   1095
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "##0.00;(##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim11 
         Height          =   285
         Left            =   2025
         TabIndex        =   143
         Top             =   2490
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim21 
         Height          =   285
         Left            =   2025
         TabIndex        =   145
         Top             =   2820
         Visible         =   0   'False
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV2"
      End
      Begin INFOHR_Controls.CodeLookup clpProvE 
         Height          =   285
         Left            =   2025
         TabIndex        =   146
         Tag             =   "30-Province Code"
         Top             =   2160
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   7250
         TabIndex        =   142
         Tag             =   "00-Supervisory Code for cheque sorting "
         Top             =   1800
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSP"
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         Height          =   285
         Left            =   7250
         TabIndex        =   144
         Tag             =   "00-Home Work Center"
         Top             =   2160
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMWC"
         MaxLength       =   12
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Tax Code WI"
         Height          =   195
         Index           =   47
         Left            =   5040
         TabIndex        =   172
         Top             =   2160
         Width           =   1620
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   46
         Left            =   5040
         TabIndex        =   171
         Top             =   1800
         Width           =   2010
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "State of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   45
         Left            =   0
         TabIndex        =   169
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblVadim11 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
         Height          =   195
         Left            =   0
         TabIndex        =   168
         Top             =   2490
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblVadim21 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 2"
         Height          =   195
         Left            =   0
         TabIndex        =   167
         Top             =   2820
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Flag 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   0
         TabIndex        =   163
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Exemptions"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   42
         Left            =   0
         TabIndex        =   157
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Tax"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   38
         Left            =   0
         TabIndex        =   156
         Top             =   1095
         Width           =   675
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Tax %"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   39
         Left            =   0
         TabIndex        =   155
         Top             =   1455
         Width           =   840
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State Extra Tax %"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   43
         Left            =   5040
         TabIndex        =   154
         Top             =   1455
         Width           =   1260
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State Extra Tax"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   44
         Left            =   5040
         TabIndex        =   153
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "State Exemptions"
         Height          =   255
         Index           =   36
         Left            =   5040
         TabIndex        =   152
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax Marital Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   0
         TabIndex        =   151
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Federal/W4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   150
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   149
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Tax Marital Status"
         Height          =   255
         Left            =   5040
         TabIndex        =   148
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frmGeneral 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   360
      TabIndex        =   104
      Top             =   4080
      Width           =   10425
      Begin VB.TextBox txtUserText1 
         Appearance      =   0  'Flat
         DataField       =   "ED_USER_TEXT1"
         DataSource      =   " "
         Height          =   280
         Left            =   6690
         MaxLength       =   20
         TabIndex        =   55
         Tag             =   "00-User Text 1"
         Top             =   2130
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CheckBox chkPenFixed 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   8040
         TabIndex        =   56
         Tag             =   "40-Fixed  - y/n"
         Top             =   1800
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.ComboBox cmbGrossCalc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   43
         Tag             =   "10-Choose C.P.P Code from list"
         Top             =   670
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbPayFreq 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6690
         TabIndex        =   50
         Tag             =   "10-Choose Pay Vacation Every Period Code from list"
         Top             =   1380
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbWCB 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6690
         TabIndex        =   38
         Tag             =   "10-Choose W.S.I.B Code"
         Top             =   -20
         Visible         =   0   'False
         Width           =   885
      End
      Begin INFOHR_Controls.CodeLookup clpVadim1 
         DataField       =   "ED_VADIM1"
         Height          =   285
         Left            =   2115
         TabIndex        =   52
         Top             =   1680
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_SUPCODE"
         Height          =   285
         Index           =   1
         Left            =   2115
         TabIndex        =   40
         Tag             =   "00-Supervisory Code for cheque sorting "
         Top             =   360
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSP"
      End
      Begin INFOHR_Controls.CodeLookup clpProv 
         DataField       =   "ED_PROVEMP"
         Height          =   285
         Left            =   2115
         TabIndex        =   37
         Tag             =   "30-Province Code"
         Top             =   0
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin VB.TextBox txtGrossCalc 
         Appearance      =   0  'Flat
         DataField       =   "ED_GROSSCD"
         Height          =   285
         Left            =   2430
         MaxLength       =   1
         TabIndex        =   42
         Tag             =   "00-Income Tax Applicable- Y/N"
         Top             =   690
         Width           =   735
      End
      Begin VB.TextBox txtGarn 
         Appearance      =   0  'Flat
         DataField       =   "ED_GARN"
         Height          =   285
         Left            =   2430
         MaxLength       =   11
         TabIndex        =   46
         Tag             =   "20-Enter garnishee $ or %"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtWCB 
         Appearance      =   0  'Flat
         DataField       =   "ED_WCB"
         Height          =   285
         Left            =   6690
         MaxLength       =   1
         TabIndex        =   39
         Tag             =   "00-Workers Compensation Board "
         Top             =   0
         Width           =   375
      End
      Begin VB.ComboBox cmbPenCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6690
         TabIndex        =   41
         Tag             =   "10-Choose Pension Code"
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtPension 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_PENSION"
         Height          =   285
         Left            =   7470
         MaxLength       =   1
         TabIndex        =   107
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cmbCPP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6720
         TabIndex        =   44
         Tag             =   "10-Choose C.P.P Code from list"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtCPP 
         Appearance      =   0  'Flat
         DataField       =   "ED_CPP"
         Height          =   285
         Left            =   7440
         MaxLength       =   1
         TabIndex        =   45
         Tag             =   "01-Canadian Pension Plan Reference"
         Top             =   690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPAYFREQ 
         Appearance      =   0  'Flat
         DataField       =   "ED_PAYFREQ"
         Height          =   285
         Left            =   6690
         MaxLength       =   1
         TabIndex        =   51
         Tag             =   "00-Pay Frequency"
         Top             =   1395
         Width           =   765
      End
      Begin MSMask.MaskEdBox medCHDSUP 
         DataField       =   "ED_CHDSUP"
         Height          =   285
         Left            =   2430
         TabIndex        =   49
         Tag             =   "10-Federal Alimony Child Support"
         Top             =   1350
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "#####"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim2 
         DataField       =   "ED_VADIM2"
         Height          =   285
         Left            =   2115
         TabIndex        =   53
         Top             =   2010
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV2"
      End
      Begin MSMask.MaskEdBox medPenPct 
         DataField       =   "ED_PENPCT"
         Height          =   285
         Left            =   6690
         TabIndex        =   54
         Tag             =   "10-Enter Pension Percentage "
         Top             =   1770
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox medVadim1 
         Height          =   285
         Left            =   9360
         TabIndex        =   161
         Tag             =   "10-Enter Vacation Pay Percentage "
         Top             =   2400
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox medVadim2 
         Height          =   285
         Left            =   8520
         TabIndex        =   162
         Tag             =   "10-Enter Vacation Pay Percentage "
         Top             =   2400
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin VB.ComboBox cmbWSIBCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6690
         TabIndex        =   47
         Tag             =   "Workplace Safety and Ins. Board of Ontario"
         Top             =   1050
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.TextBox txtWSIBCde 
         Appearance      =   0  'Flat
         DataField       =   "ED_WCBCODE"
         Height          =   285
         Left            =   6690
         MaxLength       =   6
         TabIndex        =   48
         Tag             =   "00-Workplace Safety and Ins. Board of Ontario"
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label lblUserText1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Text 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   173
         Top             =   2130
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblPPercFixed 
         AutoSize        =   -1  'True
         Caption         =   "Fixed"
         Height          =   195
         Left            =   7560
         TabIndex        =   170
         Top             =   1800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pension Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   5040
         TabIndex        =   130
         Top             =   1815
         Width           =   1440
      End
      Begin VB.Label lblVadim2 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 2"
         Height          =   195
         Left            =   0
         TabIndex        =   129
         Top             =   2070
         Width           =   945
      End
      Begin VB.Label lblVadim1 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
         Height          =   195
         Left            =   0
         TabIndex        =   128
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   116
         Top             =   15
         Width           =   1710
      End
      Begin VB.Label lblSupervisor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   115
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblCalcCode 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax Applicable"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   114
         Top             =   690
         Width           =   2055
      End
      Begin VB.Label lblGarn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Garnishee"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   113
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pension Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   5040
         TabIndex        =   112
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C.P.P."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   5025
         TabIndex        =   111
         Top             =   735
         Width           =   450
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WSIB Code"
         Height          =   195
         Index           =   35
         Left            =   5010
         TabIndex        =   110
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Federal Alimony Child Support"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   0
         TabIndex        =   109
         Top             =   1395
         Width           =   2100
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Frequency"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5025
         TabIndex        =   108
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "W.S.I.B."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   5020
         TabIndex        =   106
         Top             =   30
         Width           =   600
      End
   End
   Begin VB.ComboBox cboDepositCode 
      Height          =   315
      Left            =   960
      TabIndex        =   160
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboDepositCode2 
      Height          =   315
      Left            =   960
      TabIndex        =   159
      Top             =   1500
      Width           =   1095
   End
   Begin VB.ComboBox cboDepositCode3 
      Height          =   315
      Left            =   960
      TabIndex        =   158
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame fraOUTAddr 
      Caption         =   "Out of Province Address"
      Height          =   1635
      Left            =   10800
      TabIndex        =   120
      Top             =   5010
      Visible         =   0   'False
      Width           =   405
      Begin VB.ComboBox comOUTCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6060
         TabIndex        =   90
         Tag             =   "00-Country"
         Top             =   960
         Width           =   1320
      End
      Begin VB.TextBox txtOUTCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_OUTCOUNTRY"
         Height          =   255
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   126
         Tag             =   "01-Country"
         Top             =   720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkOUTADDRT4 
         Alignment       =   1  'Right Justify
         Caption         =   "Use this address for T4"
         DataField       =   "ED_OUTADDRT4"
         Height          =   195
         Left            =   390
         TabIndex        =   91
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtOUTAddr 
         Appearance      =   0  'Flat
         DataField       =   "ED_OUTADDR"
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   86
         Tag             =   "01-First Line in Address"
         Top             =   330
         Width           =   4180
      End
      Begin VB.TextBox txtOUTCity 
         Appearance      =   0  'Flat
         DataField       =   "ED_OUTCITY"
         Height          =   285
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   87
         Tag             =   "01-City"
         Top             =   645
         Width           =   1455
      End
      Begin MSMask.MaskEdBox medOUTPCode 
         DataField       =   "ED_OUTPCODE"
         Height          =   285
         Left            =   2400
         TabIndex        =   89
         Tag             =   "40-Postal Code"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "?#? #?#"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpOUTProv 
         DataField       =   "ED_OUTPROV"
         Height          =   285
         Left            =   5750
         TabIndex        =   88
         Tag             =   "30-Province Code"
         Top             =   630
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   4980
         TabIndex        =   127
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lblOUTProvDesc 
         Height          =   135
         Left            =   7800
         TabIndex        =   125
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   420
         TabIndex        =   124
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   420
         TabIndex        =   123
         Top             =   675
         Width           =   330
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   4950
         TabIndex        =   122
         Top             =   675
         Width           =   765
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   420
         TabIndex        =   121
         Top             =   990
         Width           =   1035
      End
   End
   Begin VB.TextBox txtUIC 
      Appearance      =   0  'Flat
      DataField       =   "ED_UIC"
      Height          =   285
      Left            =   8130
      MaxLength       =   2
      TabIndex        =   36
      Tag             =   "00-Unemployment Insurance Reference"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   85
      Top             =   7830
      Width           =   12840
      _Version        =   65536
      _ExtentX        =   22648
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
         Left            =   7920
         Top             =   0
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
   Begin VB.TextBox txtProvCode 
      Appearance      =   0  'Flat
      DataField       =   "ED_ProvCode"
      Height          =   285
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   31
      Tag             =   "00- Provincial Code"
      Top             =   2940
      Width           =   855
   End
   Begin VB.CheckBox chkProvForm 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1620
      TabIndex        =   29
      Tag             =   "40-Provincial Form  - y/n"
      Top             =   3000
      Width           =   225
   End
   Begin VB.TextBox txtExtAmt 
      Appearance      =   0  'Flat
      DataField       =   "ED_EXTAMT"
      Height          =   285
      Left            =   8160
      MaxLength       =   6
      TabIndex        =   28
      Tag             =   "10-Exempt Amount"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtFedTax 
      Appearance      =   0  'Flat
      DataField       =   "ED_FEDTAX"
      Height          =   285
      Left            =   8160
      MaxLength       =   2
      TabIndex        =   25
      Tag             =   "10-Federal Tax Method"
      Top             =   2220
      Width           =   855
   End
   Begin VB.CheckBox chkDirectDeposit 
      Alignment       =   1  'Right Justify
      Caption         =   "   "
      CausesValidation=   0   'False
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Tag             =   "40-Direct Deposit - y/n"
      Top             =   690
      Width           =   195
   End
   Begin VB.ComboBox cmbUIC 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7050
      TabIndex        =   35
      Tag             =   "10-Choose E.I. Code from list"
      Top             =   3720
      Width           =   885
   End
   Begin VB.TextBox txtTD1Code 
      Appearance      =   0  'Flat
      DataField       =   "ED_TD1CODE"
      Height          =   285
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   24
      Tag             =   "00-TD1 code as reported on the TD1 Form"
      Top             =   2220
      Width           =   855
   End
   Begin VB.TextBox txtAccount3 
      Appearance      =   0  'Flat
      DataField       =   "ED_ACCOUNT3"
      Height          =   285
      Left            =   4305
      MaxLength       =   25
      TabIndex        =   19
      Tag             =   "00-Bank Account Reference - only if Direct Deposit"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtAccount2 
      Appearance      =   0  'Flat
      DataField       =   "ED_ACCOUNT2"
      Height          =   285
      Left            =   4305
      MaxLength       =   25
      TabIndex        =   12
      Tag             =   "00-Bank Account Reference - only if Direct Deposit"
      Top             =   1500
      Width           =   1815
   End
   Begin VB.TextBox txtAccount 
      Appearance      =   0  'Flat
      DataField       =   "ED_ACCOUNT"
      Height          =   285
      Left            =   4305
      MaxLength       =   25
      TabIndex        =   5
      Tag             =   "00-Bank Account Reference - only if Direct Deposit"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtBranchCode3 
      Appearance      =   0  'Flat
      DataField       =   "ED_BRANCH3"
      Height          =   285
      Left            =   3165
      MaxLength       =   5
      TabIndex        =   18
      Tag             =   "10-Branch Reference - only if Direct Deposit"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtBranchCode2 
      Appearance      =   0  'Flat
      DataField       =   "ED_BRANCH2"
      Height          =   285
      Left            =   3165
      MaxLength       =   5
      TabIndex        =   11
      Tag             =   "10-Branch Reference - only if Direct Deposit"
      Top             =   1500
      Width           =   735
   End
   Begin VB.TextBox txtBranchCode 
      Appearance      =   0  'Flat
      DataField       =   "ED_BRANCH"
      Height          =   285
      Left            =   3165
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "10-Branch Reference - only if Direct Deposit"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtBankCode3 
      Appearance      =   0  'Flat
      DataField       =   "ED_BANK3"
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "10-Bank Reference - only if Direct Deposit"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtBankCode2 
      Appearance      =   0  'Flat
      DataField       =   "ED_BANK2"
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   9
      Tag             =   "10-Bank Reference - only if Direct Deposit"
      Top             =   1500
      Width           =   735
   End
   Begin VB.TextBox txtBankCode 
      Appearance      =   0  'Flat
      DataField       =   "ED_BANK"
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "10-Bank Reference - only if Direct Deposit"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDepositCode3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "ED_DEPOSIT3"
      Height          =   285
      Left            =   1250
      MaxLength       =   2
      TabIndex        =   15
      Tag             =   "00-Deposit Code - only if Direct Deposit"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtDepositCode2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "ED_DEPOSIT2"
      Height          =   285
      Left            =   1250
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "00-Deposit Code - only if Direct Deposit"
      Top             =   1500
      Width           =   375
   End
   Begin VB.TextBox txtDepositCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "ED_DEPOSIT"
      Height          =   285
      Left            =   1250
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "00-Deposit Code - only if Direct Deposit"
      Top             =   1200
      Width           =   375
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   12840
      _Version        =   65536
      _ExtentX        =   22648
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
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_SURNAME"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_FNAME"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         MaxLength       =   25
         TabIndex        =   61
         Text            =   "Text6"
         Top             =   120
         Visible         =   0   'False
         Width           =   990
      End
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
         Left            =   7080
         TabIndex        =   164
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
         Caption         =   "Employee"
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
         Height          =   330
         Left            =   1320
         TabIndex        =   100
         Top             =   135
         Width           =   1425
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         DataField       =   "ED_EMPNBR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   4800
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         TabIndex        =   63
         Top             =   135
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medAmountDeposit 
      DataField       =   "ED_AMTDEPOSIT"
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Tag             =   "20-Amount to be deposited"
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medPCDeposit 
      DataField       =   "ED_PCDEPOSIT"
      Height          =   285
      Left            =   8265
      TabIndex        =   7
      Tag             =   "10-% to be deposited"
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmountDeposit2 
      DataField       =   "ED_AMTDEPOSIT2"
      Height          =   285
      Left            =   6360
      TabIndex        =   13
      Tag             =   "20-Amount to be deposited"
      Top             =   1500
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medPCDeposit2 
      DataField       =   "ED_PCDEPOSIT2"
      Height          =   285
      Left            =   8265
      TabIndex        =   14
      Tag             =   "10-% to be deposited"
      Top             =   1500
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmountDeposit3 
      DataField       =   "ED_AMTDEPOSIT3"
      Height          =   285
      Left            =   6360
      TabIndex        =   20
      Tag             =   "20-Amount to be deposited"
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medPCDeposit3 
      DataField       =   "ED_PCDEPOSIT3"
      Height          =   285
      Left            =   8265
      TabIndex        =   21
      Tag             =   "10-% to be deposited"
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTD1Amnt 
      DataField       =   "ED_TD1DOL"
      Height          =   285
      Left            =   3240
      TabIndex        =   23
      Tag             =   "20-Amount as found on TD1"
      Top             =   2160
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medTD3 
      DataField       =   "ED_TD3"
      Height          =   285
      Left            =   3240
      TabIndex        =   26
      Tag             =   "20-Extra Tax on TD1 Form"
      Top             =   2490
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medVacPPct 
      DataField       =   "ED_VACPC"
      Height          =   285
      Left            =   2790
      TabIndex        =   34
      Tag             =   "10-Enter Vacation Pay Percentage "
      Top             =   3720
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
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
   Begin MSMask.MaskEdBox medTD3PC 
      DataField       =   "ED_TD3PC"
      Height          =   285
      Left            =   5640
      TabIndex        =   27
      Tag             =   "10-Extra Tax Percentage on TD1 Form"
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medProvAmt 
      DataField       =   "ED_PROVAMT"
      Height          =   285
      Left            =   3240
      TabIndex        =   30
      Tag             =   "20-Provincial Amount"
      Top             =   2940
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox MedExtraTax 
      DataField       =   "ED_ExtraTax"
      Height          =   285
      Left            =   3240
      TabIndex        =   32
      Tag             =   "20-Extra Tax on Provincial Form"
      Top             =   3240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
   Begin MSMask.MaskEdBox medExtraTaxPC 
      DataField       =   "ED_ExtraTaxPC"
      Height          =   285
      Left            =   5640
      TabIndex        =   33
      Tag             =   "10-Extra Tax Percentage on Provincial Form"
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox chkTD1Form 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   1620
      TabIndex        =   22
      Tag             =   "40-TD1 Form  - y/n"
      Top             =   2280
      Width           =   225
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ED_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   25
      TabIndex        =   82
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ED_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   83
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ED_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3480
      MaxLength       =   25
      TabIndex        =   84
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Frame frmLinamar 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   360
      TabIndex        =   105
      Top             =   4080
      Visible         =   0   'False
      Width           =   8295
      Begin VB.ComboBox cmbToPayroll 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6700
         TabIndex        =   58
         Tag             =   "00-Choose To Payroll"
         Top             =   120
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtToPayroll 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         MaxLength       =   3
         TabIndex        =   165
         Tag             =   "00-Unemployment Insurance Reference"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkQTBTORRSP 
         Height          =   195
         Left            =   2430
         TabIndex        =   59
         Tag             =   "40-Quarterly Bonus to RRSP - y/n"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtExtrAnn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2430
         MaxLength       =   5
         TabIndex        =   57
         Tag             =   "20-Extra Annual"
         Top             =   30
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To Payroll"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   41
         Left            =   5020
         TabIndex        =   166
         Top             =   150
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblQTBTORRSP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   119
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Annual"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   0
         TabIndex        =   118
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quarterly Bonus to RRSP"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   0
         TabIndex        =   117
         Top             =   495
         Width           =   1800
      End
   End
   Begin VB.TextBox txtTransitABA3 
      Appearance      =   0  'Flat
      DataField       =   "ED_TRANSITABA3"
      Height          =   285
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   17
      Tag             =   "10-Transit/ABA - only if Direct Deposit"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.TextBox txtTransitABA2 
      Appearance      =   0  'Flat
      DataField       =   "ED_TRANSITABA2"
      Height          =   285
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "10-Transit/ABA - only if Direct Deposit"
      Top             =   1500
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.TextBox txtTransitABA 
      Appearance      =   0  'Flat
      DataField       =   "ED_TRANSITABA"
      Height          =   285
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "10-Transit/ABA - only if Direct Deposit"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Left            =   5010
      Top             =   7530
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "TD1 Form    "
      Height          =   195
      Left            =   360
      TabIndex        =   102
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Provincial Form"
      Height          =   195
      Left            =   330
      TabIndex        =   103
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Direct Deposit   "
      Height          =   195
      Left            =   330
      TabIndex        =   101
      Top             =   690
      Width           =   1140
   End
   Begin VB.Label lblProvForm 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ED_ProvForm"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1200
      TabIndex        =   99
      Top             =   3330
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   2160
      TabIndex        =   98
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   4740
      TabIndex        =   97
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   2340
      TabIndex        =   96
      Top             =   3300
      Width           =   795
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax %"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   95
      Top             =   3300
      Width           =   840
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax Method"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   6600
      TabIndex        =   94
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   22
      Left            =   6960
      TabIndex        =   93
      Top             =   2550
      Width           =   1110
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax %"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   4680
      TabIndex        =   92
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E.I. Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   5380
      TabIndex        =   81
      Top             =   3750
      Width           =   660
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Pay Percentage"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   80
      Top             =   3780
      Width           =   2295
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Tax"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   2340
      TabIndex        =   79
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TD1 Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4800
      TabIndex        =   78
      Top             =   2250
      Width           =   735
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TD1 Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   2220
      TabIndex        =   77
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label lblTD1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ED_TD1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      TabIndex        =   76
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "% Deposited"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   8265
      TabIndex        =   75
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Deposited"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   6345
      TabIndex        =   74
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   4305
      TabIndex        =   73
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   3165
      TabIndex        =   72
      Top             =   960
      Width           =   930
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2085
      TabIndex        =   71
      Top             =   960
      Width           =   795
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   645
      TabIndex        =   70
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank 3"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   330
      TabIndex        =   69
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   330
      TabIndex        =   68
      Top             =   1500
      Width           =   510
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   330
      TabIndex        =   67
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lblDirectDeposit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ED_DDI"
      Enabled         =   0   'False
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
      Height          =   255
      Left            =   2160
      TabIndex        =   66
      Top             =   620
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmEBANK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OWCB, OPENSION, OVACPC
Dim OWSIBCDE 'Jaddy 11/22/99
Dim OTD1CODE, OTD1DOL, OTD3, OTD1, OSUPCODE, ODDI, oProvEmp
Dim OUIC, OCPP
Dim OGROSCALC, OGARN   'Laura nov 11, 1997
Dim ODEPCODE, OBRANCH, OBANK, OACCOUNT, OAMTDEPOSIT, OPCDEPOSIT 'ADDED BY RAUBREY 6/17/97
Dim ODEPCODE2, OBRANCH2, OBANK2, OACCOUNT2, OAMTDEPOSIT2, OPCDEPOSIT2 'ADDED BY RAUBREY 6/17/97
Dim ODEPCODE3, OBRANCH3, OBANK3, OACCOUNT3, OAMTDEPOSIT3, OPCDEPOSIT3 'ADDED BY RAUBREY 6/17/97
Dim OTRANSITABA, OTRANSITABA2, OTRANSITABA3
Dim oFedTax, oExtAmt, oProvForm, oProvAmt, oExtraTax, oExtraTaxPC, oProvCode, oProvCode2, oFedAliChd, oStateExtraTax, oStateExtraTaxPC
Dim oPAYFREQ
Dim OExtrAnn, OQTBTORRSP
Dim oVadim1, OVadim2, OVadim11, OVadim21, OSUPCODE2
Dim OOUTADDR, OOUTCITY, OOUTPROV, OOUTCOUNTRY, OOUTPCODE, OOUTADDRT4
Dim OPenPct, oHOMEWRKCNT ' oCOMBINATION
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim fglbNew As Integer
Dim GLfocus, etGLfocus, etGLfocus2
Dim SPCEICode
Dim doOnce As Boolean

Private Function AUDITBANK()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xDiv, xPT
Dim xBatchID
Dim xclpVadim1 As Double
Dim xOVadim1, xOVadim2 As Double
Dim x
On Error GoTo AUDIT_ERR
AUDITBANK = False

'New Hire makes using * worthwhile

rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

Dim UpdateAudit As Boolean
Dim Banks As New Collection
If isChanged_Bank(Banks, OBANK, txtBankCode) Then UpdateAudit = True
If isChanged_Bank(Banks, OBRANCH, txtBranchCode) Then UpdateAudit = True
If isChanged_Bank(Banks, OACCOUNT, txtAccount) Then UpdateAudit = True
If isChanged_Bank(Banks, OAMTDEPOSIT, medAmountDeposit, True) Then UpdateAudit = True
If isChanged_Bank(Banks, OPCDEPOSIT, medPCDeposit, True) Then UpdateAudit = True
If isChanged_Bank(Banks, ODEPCODE, txtDepositCode) Then UpdateAudit = True

'Town of Greater Napanee - Ticket #24375 - They do allow more than 1 account (glbCompSerial = "S/N - 2447W")
'Ticket #23795 - Town of LaSalle - Not Town of Lasalle
'Town of Aurora - Ticket #20931 - as per mapping document - do not transfer Bank 2 and Bank 3
If glbCompSerial <> "S/N - 2378W" And glbCompSerial <> "S/N - 2379W" Then
    If isChanged_Bank(Banks, OBANK2, txtBankCode2) Then UpdateAudit = True
    If isChanged_Bank(Banks, OBRANCH2, txtBranchCode2) Then UpdateAudit = True
    If isChanged_Bank(Banks, OACCOUNT2, txtAccount2) Then UpdateAudit = True
    If isChanged_Bank(Banks, OAMTDEPOSIT2, medAmountDeposit2, True) Then UpdateAudit = True
    If isChanged_Bank(Banks, OPCDEPOSIT2, medPCDeposit2, True) Then UpdateAudit = True
    If isChanged_Bank(Banks, ODEPCODE2, txtDepositCode2) Then UpdateAudit = True
    
    If isChanged_Bank(Banks, OBANK3, txtBankCode3) Then UpdateAudit = True
    If isChanged_Bank(Banks, OBRANCH3, txtBranchCode3) Then UpdateAudit = True
    If isChanged_Bank(Banks, OACCOUNT3, txtAccount3) Then UpdateAudit = True
    If isChanged_Bank(Banks, OAMTDEPOSIT3, medAmountDeposit3, True) Then UpdateAudit = True
    If isChanged_Bank(Banks, OPCDEPOSIT3, medPCDeposit3, True) Then UpdateAudit = True
    If isChanged_Bank(Banks, ODEPCODE3, txtDepositCode3) Then UpdateAudit = True
End If

If UpdateAudit Then Call Passing_Bank_Changes(Banks, glbLEE_ID)

Dim HRChanges As New Collection
'Ticket #28990 - City of Campbell River don't want Vacation Pay % to transfer to Vadim
If glbCompSerial <> "S/N - 2458W" Then
    If isChanged_Field(HRChanges, OVACPC, medVacPPct) Then UpdateAudit = True
End If
If isChanged_Field(HRChanges, OWCB, txtWCB) Then UpdateAudit = True
If isChanged_Field(HRChanges, OPENSION, txtPension) Then UpdateAudit = True
If isChanged_Field(HRChanges, OWSIBCDE, txtWSIBCde) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1CODE, txtTD1Code) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1DOL, medTD1Amnt, True) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD3, medTD3, True) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1, lblTD1) Then UpdateAudit = True
If isChanged_Field(HRChanges, OSUPCODE, clpCode(1)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODDI, lblDirectDeposit) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvEmp, clpProv) Then UpdateAudit = True

If glbWFC And fraUSA.Visible Then
    If isChanged_Field(HRChanges, OUIC, txtFedExemp) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OCPP, txtStateExemption) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OGROSCALC, txtFedMarry) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTax, medFedExtra, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTaxPC, medFedExtraPC, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OVadim11, clpVadim11) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OVadim21, clpVadim21) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oProvCode2, clpProvE) Then UpdateAudit = True
    'Ticket #22553 Franks 09/18/2012
    If isChanged_Field(HRChanges, OSUPCODE2, clpCode(0)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oHOMEWRKCNT, clpHOME) Then UpdateAudit = True
    'Ticket #23747 Franks 05/27/2013
    If isChanged_Field(HRChanges, oStateExtraTax, medStateExtra, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oStateExtraTaxPC, medStateExtraPC, True) Then UpdateAudit = True
Else
    If isChanged_Field(HRChanges, OUIC, txtUIC) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OCPP, txtCPP) Then UpdateAudit = True
    'Ticket #25469 - City of Campbell River - Transfer after checking Prov Amount. The order of transfer is creating an
    'issue when both Prov Amt and Income Tax Applicable is turned ON
    'Not now
    If glbCompSerial <> "S/N - 2458W" Then
        If isChanged_Field(HRChanges, OGROSCALC, txtGrossCalc) Then UpdateAudit = True
    End If
    If isChanged_Field(HRChanges, oExtraTax, MedExtraTax, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTaxPC, medExtraTaxPC, True) Then UpdateAudit = True
End If
If isChanged_Field(HRChanges, OGARN, txtGarn) Then UpdateAudit = True
If isChanged_Field(HRChanges, oFedTax, txtFedTax) Then UpdateAudit = True
If isChanged_Field(HRChanges, oExtAmt, txtExtAmt) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvForm, lblProvForm) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvAmt, medProvAmt, True) Then UpdateAudit = True

    'Ticket #25469 - City of Campbell River - Transfer after checking Prov Amount. The order of transfer is creating an
    'issue when both Prov Amt and Income Tax Applicable is turned ON
    'Transfer/check now
    If glbCompSerial = "S/N - 2458W" Then
        If isChanged_Field(HRChanges, OGROSCALC, txtGrossCalc) Then UpdateAudit = True
    End If

If isChanged_Field(HRChanges, oProvCode, txtProvCode) Then UpdateAudit = True
If isChanged_Field(HRChanges, OExtrAnn, txtExtrAnn) Then UpdateAudit = True
If isChanged_Field(HRChanges, OQTBTORRSP, lblQTBTORRSP) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTADDR, txtOUTAddr) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTCITY, txtOUTCity) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTPROV, clpOUTProv) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTCOUNTRY, comOUTCountry) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTPCODE, medOUTPCode) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTADDRT4, chkOUTADDRT4) Then UpdateAudit = True

If isChanged_Field(HRChanges, OTRANSITABA, txtTransitABA) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTRANSITABA2, txtTransitABA2) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTRANSITABA3, txtTransitABA3) Then UpdateAudit = True
If isChanged_Field(HRChanges, OPenPct, medPenPct) Then UpdateAudit = True

If glbCompSerial = "S/N - 2276W" Then    'City of Niagara Falls
    'Federal Alimony Child Support
    If isChanged_Field(HRChanges, oFedAliChd, medCHDSUP) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oPAYFREQ, txtPAYFREQ) Then UpdateAudit = True
End If

'Vadim Field 1
'For City of Timmins or City of Niagara Falls
If (glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W") And clpVadim1 <> "" Then
    medVadim1.DataField = "ED_VADIM1"
    medVadim1 = (Val(clpVadim1) / 100)
    xOVadim1 = (Val(oVadim1) / 100)
    If isChanged_Field(HRChanges, xOVadim1, medVadim1, True) Then UpdateAudit = True
    medVadim1.DataField = ""
Else
    If glbCompSerial = "S/N - 2363W" Then    'City of Kawartha Lakes
        If isChanged_Field(HRChanges, oVadim1, clpVadim1, True) Then UpdateAudit = True
    Else
        'Town of Lasalle - Do not transfer Vadim 1
        'Ticket #24996 - City of Campbell River - Do not transfer Vadim 1
        If glbCompSerial <> "S/N - 2379W" And glbCompSerial <> "S/N - 2458W" Then
            If isChanged_Field(HRChanges, oVadim1, clpVadim1) Then UpdateAudit = True
        End If
    End If
End If

'Vadim Field 2
'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then 'Or glbCompSerial = "S/N - 2375W" Then
    If clpVadim2 <> "" Then
        medVadim2.DataField = "ED_VADIM2"
        medVadim2 = (Val(clpVadim2) / 100)
        xOVadim2 = (Val(OVadim2) / 100)
        If isChanged_Field(HRChanges, xOVadim2, medVadim2, True) Then UpdateAudit = True
        medVadim2.DataField = ""
    Else
        xOVadim2 = (Val(OVadim2) / 100)
        If isChanged_Field(HRChanges, xOVadim2, clpVadim2, True) Then UpdateAudit = True
    End If
Else
    'Town of Lasalle - Do not transfer Vadim 2
    'Ticket #24996 - City of Campbell River - Do not transfer Vadim 2
    If glbCompSerial <> "S/N - 2379W" And glbCompSerial <> "S/N - 2458W" Then
        If isChanged_Field(HRChanges, OVadim2, clpVadim2) Then UpdateAudit = True
    End If
End If

Call Passing_Changes(HRChanges, Banking, "M", Date, glbLEE_ID)

If UpdateAudit Then
    GoTo MODUPD 'ticket# 7298
Else
    GoTo MODNOUPD
End If

MODUPD:
    rsTB.Open "select ED_DIV,ED_PT FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    
    If Not rsTB.EOF Then
      'xDiv = rsTB("ED_DIV")
      'xPT = rsTB("ED_PT")
        If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
    Else
      xDiv = ""
      xPT = ""
    End If
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    
    If glbPayWeb Or glbInsync Then
        rsTA("AU_WCBCODE") = txtWSIBCde
        rsTA("AU_WCB") = txtWCB
        rsTA("AU_PENSION") = txtPension
        rsTA("AU_DDI") = lblDirectDeposit
        rsTA("AU_DEPOSIT") = txtDepositCode
        rsTA("AU_BRANCH") = txtBranchCode
        rsTA("AU_BANK") = txtBankCode
        rsTA("AU_ACCOUNT") = txtAccount
        If IsNumeric(medAmountDeposit) Then rsTA("AU_AMTDEPOSIT") = Val(medAmountDeposit)
        If IsNumeric(medPCDeposit) Then rsTA("AU_PCDEPOSIT") = Val(medPCDeposit)
        rsTA("AU_DEPOSIT2") = txtDepositCode2
        rsTA("AU_BRANCH2") = txtBranchCode2
        rsTA("AU_BANK2") = txtBankCode2
        rsTA("AU_ACCOUNT2") = txtAccount2
        If IsNumeric(medAmountDeposit2) Then rsTA("AU_AMTDEPOSIT2") = Val(medAmountDeposit2)
        If IsNumeric(medPCDeposit2) Then rsTA("AU_PCDEPOSIT2") = Val(medPCDeposit2)
        rsTA("AU_DEPOSIT3") = txtDepositCode3
        rsTA("AU_BRANCH3") = txtBranchCode3
        rsTA("AU_BANK3") = txtBankCode3
        rsTA("AU_ACCOUNT3") = txtAccount3
        If IsNumeric(medAmountDeposit3) Then rsTA("AU_AMTDEPOSIT3") = Val(medAmountDeposit3)
        If IsNumeric(medPCDeposit3) Then rsTA("AU_PCDEPOSIT3") = Val(medPCDeposit3)
        If IsNumeric(medVacPPct) Then rsTA("AU_VACPC") = medVacPPct * 100
        If IsNumeric(OVACPC) Then rsTA("AU_OLDVAC") = OVACPC * 100
        rsTA("AU_TD1CODE") = txtTD1Code & ""
        If IsNumeric(medTD1Amnt) Then rsTA("AU_TD1DOL") = medTD1Amnt
        If IsNumeric(OTD1DOL) Then rsTA("AU_OLDTD1") = Val(OTD1DOL)                       'sbh added val...
        If IsNumeric(medTD3) Then rsTA("AU_TD3") = Val(medTD3)
        If IsNumeric(OTD3) Then rsTA("AU_OLDTD3") = Val(OTD3)   'sbh added...
        rsTA("AU_TD1") = lblTD1
        rsTA("AU_SUPCODE") = clpCode(1).Text
        rsTA("AU_DDI") = lblDirectDeposit
        rsTA("AU_PROVEMP") = clpProv.Text
        rsTA("AU_UIC") = txtUIC
        rsTA("AU_CPP") = txtCPP
        If IsNumeric(txtFedTax) Then rsTA("AU_FedTax") = txtFedTax
        
        rsTA("AU_ProvForm") = lblProvForm
        
        If IsNumeric(txtExtAmt) Then rsTA("AU_ExtAmt") = txtExtAmt
        If IsNumeric(medProvAmt) Then rsTA("AU_ProvAmt") = medProvAmt
        If IsNumeric(MedExtraTax) Then rsTA("AU_ExtraTax") = MedExtraTax
        If IsNumeric(medExtraTaxPC) Then rsTA("AU_ExtraTaxPC") = medExtraTaxPC
        rsTA("AU_ProvCode") = txtProvCode
        If IsNumeric(medPenPct) Then rsTA("AU_PENPCT") = medPenPct
        If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership
            If Len(clpVadim1.Text) > 0 Then rsTA("AU_VADIM1") = clpVadim1.Text
            If Len(cmbCPP.Text) > 0 Then
                rsTA("AU_CPP") = cmbCPP.Text
            End If
        End If
    ElseIf glbVadim Then
        rsTA("AU_DDI") = lblDirectDeposit
        rsTA("AU_DEPOSIT") = txtDepositCode
        rsTA("AU_BRANCH") = txtBranchCode
        rsTA("AU_BANK") = txtBankCode
        rsTA("AU_ACCOUNT") = txtAccount
        If IsNumeric(medAmountDeposit) Then rsTA("AU_AMTDEPOSIT") = Val(medAmountDeposit)
        If IsNumeric(medPCDeposit) Then rsTA("AU_PCDEPOSIT") = Val(medPCDeposit)
        rsTA("AU_DEPOSIT2") = txtDepositCode2
        rsTA("AU_BRANCH2") = txtBranchCode2
        rsTA("AU_BANK2") = txtBankCode2
        rsTA("AU_ACCOUNT2") = txtAccount2
        If IsNumeric(medAmountDeposit2) Then rsTA("AU_AMTDEPOSIT2") = Val(medAmountDeposit2)
        If IsNumeric(medPCDeposit2) Then rsTA("AU_PCDEPOSIT2") = Val(medPCDeposit2)
        rsTA("AU_DEPOSIT3") = txtDepositCode3
        rsTA("AU_BRANCH3") = txtBranchCode3
        rsTA("AU_BANK3") = txtBankCode3
        rsTA("AU_ACCOUNT3") = txtAccount3
        If IsNumeric(medAmountDeposit3) Then rsTA("AU_AMTDEPOSIT3") = Val(medAmountDeposit3)
        If IsNumeric(medPCDeposit3) Then rsTA("AU_PCDEPOSIT3") = Val(medPCDeposit3)
''        If IsNumeric(medVacPPct) Then rsTA("AU_VACPC") = medVacPPct * 100
''        If IsNumeric(OVACPC) Then rsTA("AU_OLDVAC") = OVACPC * 100
''        rsTA("AU_PROVEMP") = clpProv.Text
''        rsTA("AU_UIC") = txtUIC
''        rsTA("AU_CPP") = txtCPP
''        rsTA("AU_ProvCode") = txtProvCode
    Else
        
        If OWCB <> txtWCB Then rsTA("AU_WCB") = txtWCB
        If OWSIBCDE <> txtWSIBCde Then rsTA("AU_WCBCODE") = txtWSIBCde  'Jaddy 11/22/99
        
        If OPENSION <> txtPension Then
            If Len(txtPension) > 0 Then
                rsTA("AU_PENSION") = txtPension
            Else
                rsTA("AU_PENSION") = "-"
            End If
        End If
        If ODDI <> lblDirectDeposit Then rsTA("AU_DDI") = lblDirectDeposit
        '** Bank 1
        If ODEPCODE <> txtDepositCode Or OBRANCH <> txtBranchCode Or OBANK <> txtBankCode Or OACCOUNT <> txtAccount Or Val(OAMTDEPOSIT) <> Val(medAmountDeposit) Or Val(OPCDEPOSIT) <> Val(medPCDeposit) Then
            If Len(txtDepositCode) > 0 Then
                rsTA("AU_DEPOSIT") = txtDepositCode
            Else
                rsTA("AU_DEPOSIT") = ""
            End If
            If Len(txtBranchCode) > 0 Then
                rsTA("AU_BRANCH") = txtBranchCode
            Else
                rsTA("AU_BRANCH") = "-"
            End If
            If Len(txtBankCode) > 0 Then
                rsTA("AU_BANK") = txtBankCode
            Else
                rsTA("AU_BANK") = "-"
            End If
            If Len(txtAccount) > 0 Then
                rsTA("AU_ACCOUNT") = txtAccount
            Else
                rsTA("AU_ACCOUNT") = "-"
            End If
            If Len(medAmountDeposit) > 0 Then
                If Val(medAmountDeposit) > 0 Then
                    rsTA("AU_AMTDEPOSIT") = Val(medAmountDeposit)
                End If
            End If
            If Len(medPCDeposit) > 0 Then
                If Val(medPCDeposit) > 0 Then
                    rsTA("AU_PCDEPOSIT") = Val(medPCDeposit)
                End If
            End If
            rsTA("AU_DDI") = "Y"
            'Ticket #23747 Franks 05/27/2013
            If txtTransitABA.Visible Then
                If Len(txtTransitABA.Text) > 0 Then
                    rsTA("AU_TRANSITABA") = txtTransitABA.Text
                End If
            End If
        End If
        
        If OTRANSITABA <> txtTransitABA Then rsTA("AU_TRANSITABA") = txtTransitABA
        If OTRANSITABA2 <> txtTransitABA2 Then rsTA("AU_TRANSITABA2") = txtTransitABA2
        If OTRANSITABA3 <> txtTransitABA3 Then rsTA("AU_TRANSITABA3") = txtTransitABA3
        
        
        '~~~~~~~~~~~~~~~~ADDED BY RAUBREY 6/17/97~~~~~~~~~~~~~~~~~~~~
        
        'If Val(OAMTDEPOSIT) <> Val(medAmountDeposit) Then
        '    rsTA("AU_AMTDEPOSIT") = Val(medAmountDeposit)
        'End If
        'If Val(OPCDEPOSIT) <> Val(medPCDeposit) Then
        '    rsTA("AU_PCDEPOSIT") = Val(medPCDeposit)
        'End If
        
        '**BANK2
        '~~~~~~~~~changed by Laura 11/1/97
        If ODEPCODE2 <> txtDepositCode2 Or OBRANCH2 <> txtBranchCode2 Or OBANK2 <> txtBankCode2 Or OACCOUNT2 <> txtAccount2 Or Val(OAMTDEPOSIT2) <> Val(medAmountDeposit2) Or Val(OPCDEPOSIT2) <> Val(medPCDeposit2) Then
            If Len(txtDepositCode2) > 0 Then
                rsTA("AU_DEPOSIT2") = txtDepositCode2
            Else
                rsTA("AU_DEPOSIT2") = ""
            End If
            If Len(txtBranchCode2) > 0 Then
                rsTA("AU_BRANCH2") = txtBranchCode2
            Else
                rsTA("AU_BRANCH2") = "-"
            End If
            If Len(txtBankCode2) > 0 Then
                rsTA("AU_BANK2") = txtBankCode2
            Else
                rsTA("AU_BANK2") = "-"
            End If
            If Len(txtAccount2) > 0 Then
                rsTA("AU_ACCOUNT2") = txtAccount2
            Else
                rsTA("AU_ACCOUNT2") = "-"
            End If
            If Len(medAmountDeposit2) > 0 Then
                If Val(medAmountDeposit2) > 0 Then
                    rsTA("AU_AMTDEPOSIT2") = Val(medAmountDeposit2)
                End If
            End If
            If Len(medPCDeposit2) > 0 Then
                If Val(medPCDeposit2) > 0 Then
                    rsTA("AU_PCDEPOSIT2") = Val(medPCDeposit2)
                End If
            End If
            rsTA("AU_DDI") = "Y"
            'Ticket #23747 Franks 05/27/2013
            If txtTransitABA2.Visible Then
                If Len(txtTransitABA2.Text) > 0 Then
                    rsTA("AU_TRANSITABA2") = txtTransitABA2.Text
                End If
            End If
        End If
        
        'If Val(OAMTDEPOSIT2) <> Val(medAmountDeposit2) Then
        '    rsTA("AU_AMTDEPOSIT2") = Val(medAmountDeposit2)
        'End If
        'If Val(OPCDEPOSIT2) <> Val(medPCDeposit2) Then
        '    rsTA("AU_PCDEPOSIT2") = Val(medPCDeposit2)
        'End If
        
        '**BANK3
        '~~~~~~~~~changed by Laura 11/1/97
        If ODEPCODE3 <> txtDepositCode3 Or OBRANCH3 <> txtBranchCode3 Or OBANK3 <> txtBankCode3 Or OACCOUNT3 <> txtAccount3 Or Val(OAMTDEPOSIT3) <> Val(medAmountDeposit3) Or Val(OPCDEPOSIT3) <> Val(medPCDeposit3) Then
            If Len(txtDepositCode3) > 0 Then
                rsTA("AU_DEPOSIT3") = txtDepositCode3
            Else
                rsTA("AU_DEPOSIT3") = ""
            End If
            If Len(txtBranchCode3) > 0 Then
                rsTA("AU_BRANCH3") = txtBranchCode3
            Else
                rsTA("AU_BRANCH3") = "-"
            End If
            If Len(txtBankCode3) > 0 Then
                rsTA("AU_BANK3") = txtBankCode3
            Else
                rsTA("AU_BANK3") = "-"
            End If
            If Len(txtAccount3) > 0 Then
                rsTA("AU_ACCOUNT3") = txtAccount3
            Else
                rsTA("AU_ACCOUNT3") = "-"
            End If
            If Len(medAmountDeposit3) > 0 Then
                If Val(medAmountDeposit3) > 0 Then
                    rsTA("AU_AMTDEPOSIT3") = Val(medAmountDeposit3)
                End If
            End If
            If Len(medPCDeposit3) > 0 Then
                If Val(medPCDeposit3) > 0 Then
                    rsTA("AU_PCDEPOSIT3") = Val(medPCDeposit3)
                End If
            End If
            rsTA("AU_DDI") = "Y"
            'Ticket #23747 Franks 05/27/2013
            If txtTransitABA3.Visible Then
                If Len(txtTransitABA3.Text) > 0 Then
                    rsTA("AU_TRANSITABA3") = txtTransitABA3.Text
                End If
            End If
        End If
    
        If glbCompSerial = "S/N - 2347W" Then 'Surrey Place - Audit all Banks
            If ODEPCODE <> txtDepositCode Or OBRANCH <> txtBranchCode Or OBANK <> txtBankCode Or OACCOUNT <> txtAccount Or Val(OAMTDEPOSIT) <> Val(medAmountDeposit) Or Val(OPCDEPOSIT) <> Val(medPCDeposit) Or _
                ODEPCODE2 <> txtDepositCode2 Or OBRANCH2 <> txtBranchCode2 Or OBANK2 <> txtBankCode2 Or OACCOUNT2 <> txtAccount2 Or Val(OAMTDEPOSIT2) <> Val(medAmountDeposit2) Or Val(OPCDEPOSIT2) <> Val(medPCDeposit2) Or _
                ODEPCODE3 <> txtDepositCode3 Or OBRANCH3 <> txtBranchCode3 Or OBANK3 <> txtBankCode3 Or OACCOUNT3 <> txtAccount3 Or Val(OAMTDEPOSIT3) <> Val(medAmountDeposit3) Or Val(OPCDEPOSIT3) <> Val(medPCDeposit3) Then
                
                'Bank 1
                If Len(txtDepositCode) > 0 Then
                    rsTA("AU_DEPOSIT") = txtDepositCode
                Else
                    rsTA("AU_DEPOSIT") = ""
                End If
                If Len(txtBranchCode) > 0 Then
                    rsTA("AU_BRANCH") = txtBranchCode
                Else
                    rsTA("AU_BRANCH") = "-"
                End If
                If Len(txtBankCode) > 0 Then
                    rsTA("AU_BANK") = txtBankCode
                Else
                    rsTA("AU_BANK") = "-"
                End If
                If Len(txtAccount) > 0 Then
                    rsTA("AU_ACCOUNT") = txtAccount
                Else
                    rsTA("AU_ACCOUNT") = "-"
                End If
                If Len(medAmountDeposit) > 0 Then
                    If Val(medAmountDeposit) > 0 Then
                        rsTA("AU_AMTDEPOSIT") = Val(medAmountDeposit)
                    End If
                End If
                If Len(medPCDeposit) > 0 Then
                    If Val(medPCDeposit) > 0 Then
                        rsTA("AU_PCDEPOSIT") = Val(medPCDeposit)
                    End If
                End If
                
                'Bank 2
                If Len(txtDepositCode2) > 0 Then
                    rsTA("AU_DEPOSIT2") = txtDepositCode2
                Else
                    rsTA("AU_DEPOSIT2") = ""
                End If
                If Len(txtBranchCode2) > 0 Then
                    rsTA("AU_BRANCH2") = txtBranchCode2
                Else
                    rsTA("AU_BRANCH2") = "-"
                End If
                If Len(txtBankCode2) > 0 Then
                    rsTA("AU_BANK2") = txtBankCode2
                Else
                    rsTA("AU_BANK2") = "-"
                End If
                If Len(txtAccount2) > 0 Then
                    rsTA("AU_ACCOUNT2") = txtAccount2
                Else
                    rsTA("AU_ACCOUNT2") = "-"
                End If
                If Len(medAmountDeposit2) > 0 Then
                    If Val(medAmountDeposit2) > 0 Then
                        rsTA("AU_AMTDEPOSIT2") = Val(medAmountDeposit2)
                    End If
                End If
                If Len(medPCDeposit2) > 0 Then
                    If Val(medPCDeposit2) > 0 Then
                        rsTA("AU_PCDEPOSIT2") = Val(medPCDeposit2)
                    End If
                End If
                
                'Bank 3
                If Len(txtDepositCode3) > 0 Then
                    rsTA("AU_DEPOSIT3") = txtDepositCode3
                Else
                    rsTA("AU_DEPOSIT3") = ""
                End If
                If Len(txtBranchCode3) > 0 Then
                    rsTA("AU_BRANCH3") = txtBranchCode3
                Else
                    rsTA("AU_BRANCH3") = "-"
                End If
                If Len(txtBankCode3) > 0 Then
                    rsTA("AU_BANK3") = txtBankCode3
                Else
                    rsTA("AU_BANK3") = "-"
                End If
                If Len(txtAccount3) > 0 Then
                    rsTA("AU_ACCOUNT3") = txtAccount3
                Else
                    rsTA("AU_ACCOUNT3") = "-"
                End If
                If Len(medAmountDeposit3) > 0 Then
                    If Val(medAmountDeposit3) > 0 Then
                        rsTA("AU_AMTDEPOSIT3") = Val(medAmountDeposit3)
                    End If
                End If
                If Len(medPCDeposit3) > 0 Then
                    If Val(medPCDeposit3) > 0 Then
                        rsTA("AU_PCDEPOSIT3") = Val(medPCDeposit3)
                    End If
                End If
                rsTA("AU_DDI") = "Y"
            End If
        End If
        
        'If Val(OAMTDEPOSIT3) <> Val(medAmountDeposit3) Then
        '    rsTA("AU_AMTDEPOSIT3") = Val(medAmountDeposit3)
        'End If
        'If Val(OPCDEPOSIT3) <> Val(medPCDeposit3) Then
        '    rsTA("AU_PCDEPOSIT3") = Val(medPCDeposit3)
        'End If
        
        If OVACPC <> medVacPPct Then
            If IsNumeric(medVacPPct) Then rsTA("AU_VACPC") = medVacPPct * 100
            If IsNumeric(OVACPC) Then rsTA("AU_OLDVAC") = OVACPC * 100
        End If
        
        If OTD1CODE <> txtTD1Code Then rsTA("AU_TD1CODE") = txtTD1Code & ""  'sbh added ""
        If Val(OTD1DOL) <> Val(medTD1Amnt) Then  'sbh added val...
            If IsNumeric(medTD1Amnt) Then rsTA("AU_TD1DOL") = medTD1Amnt
            If IsNumeric(OTD1DOL) Then
                rsTA("AU_OLDTD1") = Val(OTD1DOL)        'sbh added val...
            End If
        End If
        If glbWFC And fraUSA.Visible Then 'Ticket #23747 Franks 05/27/2013
            If OTD3 <> medStateExtra Then
                rsTA("AU_TD3") = Val(medStateExtra)
                rsTA("AU_OLDTD3") = Val(OTD3)  'sbh added...
            End If
        Else
            If OTD3 <> medTD3 Then
                rsTA("AU_TD3") = Val(medTD3)
                rsTA("AU_OLDTD3") = Val(OTD3)  'sbh added...
            End If
        End If
        If OTD1 <> lblTD1 Then rsTA("AU_TD1") = lblTD1
        If OSUPCODE <> clpCode(1).Text Then rsTA("AU_SUPCODE") = clpCode(1).Text
        If ODDI <> lblDirectDeposit Then rsTA("AU_DDI") = lblDirectDeposit
        If oProvEmp <> clpProv.Text Then
            If Len(clpProv.Text) > 0 Then
                rsTA("AU_PROVEMP") = clpProv.Text
            Else
                rsTA("AU_PROVEMP") = "-"
            End If
        End If
        
        'Greensboro
        If glbWFC And fraUSA.Visible Then   'And fgetSection(lblEEID) = "GREN" Then
            If OTD3 <> medStateExtra Then
                rsTA("AU_TD3") = Val(medStateExtra)
                rsTA("AU_OLDTD3") = Val(OTD3)  'sbh added...
            End If
            If (OUIC <> txtFedExemp) Or (NewHireForms.count > 0) Then      'New Hire only
                If (OUIC <> txtFedExemp) Then
                    If Not (glbCompSerial = "S/N - 2347W") Then  'Surrey Place
                        'Don't pass EI Code change to Audit for Surrey Place, Ticket# 9152
                        If Len(txtFedExemp) > 0 Then
                            rsTA("AU_UIC") = txtFedExemp
                        Else
                            rsTA("AU_UIC") = "-"
                        End If
                    End If
                Else
                    If Len(txtFedExemp) > 0 Then
                        rsTA("AU_UIC") = txtFedExemp
                    Else
                        rsTA("AU_UIC") = "-"
                    End If
                End If
            End If
            If OCPP <> txtStateExemption Then
                If Len(txtStateExemption) > 0 Then
                    rsTA("AU_CPP") = txtStateExemption
                Else
                    rsTA("AU_CPP") = "-"
                End If
            End If
            If OGROSCALC <> txtFedMarry Then rsTA("AU_GROSSCD") = txtFedMarry  'laura nov 11, 1997
            If oExtraTax <> medFedExtra Then rsTA("AU_ExtraTax") = Val(medFedExtra)
            If oExtraTaxPC <> medFedExtraPC Then rsTA("AU_ExtraTaxPC") = medFedExtraPC
            If OPENSION <> txtStatusFalg3 Then rsTA("AU_PENSION") = txtStatusFalg3
            If oStateExtraTax <> medStateExtra Then rsTA("AU_TD3") = Val(medStateExtra)
            If oStateExtraTaxPC <> medStateExtraPC Then rsTA("AU_TD1DOL") = medStateExtraPC
            If glbCompSerial = "S/N - 2217W" Then 'City of Pickering
            'Ticket #20054 Franks 04/04/2011, keep old GLNO and old Union
            Else
            If OVadim11 <> clpVadim11 Then rsTA("AU_VADIM1") = clpVadim11
            If OVadim21 <> clpVadim21 Then rsTA("AU_VADIM2") = clpVadim21
            End If
            If oProvCode2 <> clpProvE Then rsTA("AU_PROVEMP") = clpProvE
            If OSUPCODE2 <> clpCode(0) Then rsTA("AU_SUPCODE") = clpCode(0).Text
            If oHOMEWRKCNT <> clpHOME Then rsTA("AU_HOMEWRKCNT") = clpHOME.Text
        Else
            If glbWFC And fraUSA.Visible Then 'Ticket #23747 Franks 05/27/2013
                If OTD3 <> medStateExtra Then
                    rsTA("AU_TD3") = Val(medStateExtra)
                    rsTA("AU_OLDTD3") = Val(OTD3)  'sbh added...
                End If
            Else
                If OTD3 <> medTD3 Then
                    rsTA("AU_TD3") = Val(medTD3)
                    rsTA("AU_OLDTD3") = Val(OTD3)  'sbh added...
                End If
            End If
            If (OUIC <> txtUIC) Or (NewHireForms.count > 0) Then      'New Hire only
                If (OUIC <> txtUIC) Then
                    If Not (glbCompSerial = "S/N - 2347W") Then  'Surrey Place
                        'Don't pass EI Code change to Audit for Surrey Place, Ticket# 9152
                        If Len(txtUIC) > 0 Then
                            rsTA("AU_UIC") = txtUIC
                        Else
                            rsTA("AU_UIC") = "-"
                        End If
                    End If
                Else
                    If Len(txtUIC) > 0 Then
                        rsTA("AU_UIC") = txtUIC
                    Else
                        rsTA("AU_UIC") = "-"
                    End If
                End If
            End If
            If OCPP <> txtCPP Then
                If Len(txtCPP) > 0 Then
                    rsTA("AU_CPP") = txtCPP
                Else
                    rsTA("AU_CPP") = "-"
                End If
            End If
            If OGROSCALC <> txtGrossCalc Then rsTA("AU_GROSSCD") = txtGrossCalc  'laura nov 11, 1997
            If oExtraTax <> MedExtraTax Then rsTA("AU_ExtraTax") = Val(MedExtraTax)
            If oExtraTaxPC <> medExtraTaxPC Then rsTA("AU_ExtraTaxPC") = medExtraTaxPC
        End If
        
        If oFedTax <> txtFedTax Then rsTA("AU_FedTax") = txtFedTax
        If oExtAmt <> txtExtAmt Then rsTA("AU_ExtAmt") = txtExtAmt
        If oProvForm <> lblProvForm Then rsTA("AU_ProvForm") = lblProvForm
        If oProvAmt <> medProvAmt Then rsTA("AU_ProvAmt") = medProvAmt

        If oProvCode <> txtProvCode Then rsTA("AU_ProvCode") = txtProvCode
        
        
        If OGARN <> txtGarn Then rsTA("AU_GARN") = txtGarn    'laura nov 11, 1997
        If glbLinamar Then
            If OExtrAnn <> txtExtrAnn Then rsTA("AU_ExtrAnn") = txtExtrAnn
            If OQTBTORRSP <> lblQTBTORRSP Then rsTA("AU_QTBTORRSP") = lblQTBTORRSP
            If OOUTADDR <> txtOUTAddr Then rsTA("AU_OUTADDR") = txtOUTAddr
            If OOUTCITY <> txtOUTCity Then rsTA("AU_OUTCITY") = txtOUTCity
            If OOUTPROV <> clpOUTProv Then rsTA("AU_OUTPROV") = clpOUTProv
            If OOUTCOUNTRY <> comOUTCountry Then rsTA("AU_OUTCOUNTRY") = comOUTCountry
            If OOUTPCODE <> medOUTPCode Then rsTA("AU_OUTPCODE") = medOUTPCode
            If OOUTADDRT4 <> chkOUTADDRT4 Then rsTA("AU_OUTADDRT4") = chkOUTADDRT4
        End If
        If glbCompSerial = "S/N - 2217W" Then 'City of Pickering
        'Ticket #20054 Franks 04/04/2011, keep old GLNO and old Union
        Else
        If oVadim1 <> clpVadim1 Then rsTA("AU_VADIM1") = clpVadim1
        If OVadim2 <> clpVadim2 Then rsTA("AU_VADIM2") = clpVadim2
        End If
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    'Ticket #23407 Franks 03/15/2013
    'If glbCompSerial = "S/N - 2382W" Or glbWFC Then  ' Samuel - Ticket #18702 'WFC - Ticket #21543
        If IsDate(rsDATA("ED_DOH")) Then
            If CVDate(rsDATA("ED_DOH")) > Date Then
                rsTA("AU_LDATE") = rsDATA("ED_DOH")
            End If
        End If
    'End If
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    'If glbSoroc Or glbSyndesis Then
        If Not IsNull(rsDATA("ED_Payroll_ID")) Then rsTA("AU_Payroll_ID") = rsDATA("ED_Payroll_ID")
    'End If
    rsTA.Update
    If glbCompSerial = "S/N - 2347W" Then 'Surrey Place
        If Len(glbSPCTermReason) > 0 Then
            Call SPCPayrollIDAudit(rsTA)
        End If
    End If
    
    If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
        Call FamilyDayAuditSync(glbLEE_ID, rsTA)
    End If
    
MODNOUPD:
AUDITBANK = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack
Resume Next
End Function

Private Sub SPCPayrollIDAudit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsTA2 As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ
Dim rsSal As New ADODB.Recordset 'Ticket #25118 Franks 02/21/2014

    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & lblEEID & " "
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'Termination Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_SURNAME") = rsEmp("ED_SURNAME")
    rsTA("AU_FNAME") = rsEmp("ED_FNAME")
    rsTA("AU_DOT") = glbSPCTermDate
    rsTA("AU_TREAS") = glbSPCTermReason
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = lblEEID '
    rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    'Ticket #25118 Franks 02/21/2014 - begin
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT =0 AND SH_EMPNBR = " & lblEEID & " "
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsSal.EOF Then
        If Not IsNull(rsSal("SH_PAYP")) Then
            rsTA("AU_PAYP") = rsSal("SH_PAYP")
            rsTA("AU_USER_TEXT1") = "PAYROLL_TERM"
        End If
    End If
    rsSal.Close
    'Ticket #25118 Franks 02/21/2014 - end
    rsTA.Update
    
    'New Hire Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1":: rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_DIV") = rsEmp("ED_DIV") 'clpDiv.Text
    rsTA("AU_DEPTNO") = rsEmp("ED_DEPTNO")
    rsTA("AU_TITLE") = rsEmp("ED_TITLE")
    rsTA("AU_SURNAME") = rsEmp("ED_SURNAME")
    rsTA("AU_FNAME") = rsEmp("ED_FNAME")
    rsTA("AU_EMPNBR") = glbSPCNewEmpNo
    rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    rsTA("AU_ADDR1") = rsEmp("ED_ADDR1")
    rsTA("AU_ADDR2") = rsEmp("ED_ADDR2")
    rsTA("AU_CITY") = rsEmp("ED_CITY")
    rsTA("AU_PROV") = rsEmp("ED_PROV")
    rsTA("AU_COUNTRY") = rsEmp("ED_COUNTRY")
    rsTA("AU_PCODE") = rsEmp("ED_PCODE")
    rsTA("AU_PHONE") = rsEmp("ED_PHONE")
    rsTA("AU_BUSNBR") = rsEmp("ED_BUSNBR")
    rsTA("AU_DIVUPL") = rsEmp("ED_DIV")
    rsTA("AU_SEX") = rsEmp("ED_SEX")
    If Not IsNull(rsEmp("ED_SMOKER")) Then
        rsTA("AU_SMOKER") = IIf(rsEmp("ED_SMOKER"), "Yes", "No")
    End If
    rsTA("AU_DOB") = rsEmp("ED_DOB")
    rsTA("AU_DOH") = rsEmp("ED_DOH")
    rsTA("AU_SIN") = rsEmp("ED_SIN")
    rsTA("AU_DEPT_GL") = rsEmp("ED_GLNO")
    rsTA("AU_MSTAT") = rsEmp("ED_MSTAT")
    rsTA("AU_NEWEMP") = "Y"
    rsTA("AU_PTUPL") = rsEmp("ED_PT")
    rsTA("AU_LOC") = rsEmp("ED_LOC")

    rsTA("AU_ADMINBY") = rsEmp("ED_ADMINBY")
    rsTA("AU_REGION") = rsEmp("ED_REGION")
    rsTA("AU_SECTION") = rsEmp("ED_SECTION")
    rsTA("AU_HOMEOPRTNBR") = rsEmp("ED_HOMEOPRTNBR")
    rsTA("AU_HOMELINE") = rsEmp("ED_HOMELINE")
    rsTA("AU_HOMESHIFT") = rsEmp("ED_HOMESHIFT")
    rsTA("AU_HOMEWRKCNT") = rsEmp("ED_HOMEWRKCNT")
    rsTA("AU_CellPhone") = rsEmp("ED_CellPhone")
    rsTA("AU_PageNbr") = rsEmp("ED_PageNbr")
    rsTA("AU_SSN") = rsEmp("ED_SSN")
    rsTA("AU_ORG") = rsEmp("ED_ORG")

    rsTA("AU_DEPTEDATE") = rsEmp("ED_DEPTEDATE")
    rsTA("AU_DIVEDATE") = rsEmp("ED_DIVEDATE")
    rsTA("AU_DRIVERLIC") = rsEmp("ED_DRIVERLIC")
    rsTA("AU_LICPLATE1") = rsEmp("ED_LICPLATE1")
    rsTA("AU_LICPLATE2") = rsEmp("ED_LICPLATE2")
    rsTA("AU_TYPEVEHICLE") = rsEmp("ED_TYPEVEHICLE")
    rsTA("AU_PARKPERMIT1") = rsEmp("ED_PARKPERMIT1")
    rsTA("AU_PARKPERMIT2") = rsEmp("ED_PARKPERMIT2")
    rsTA("AU_BADGEID") = rsEmp("ED_BADGEID")
    rsTA("AU_MIDNAME") = rsEmp("ED_MIDNAME")
    rsTA("AU_ALIAS") = rsEmp("ED_ALIAS")
    rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA.Update
    
    '------BANK Information Begin
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    'BANK 1
    rsTA("AU_DEPOSIT") = rsEmp("ED_DEPOSIT")
    rsTA("AU_BRANCH") = rsEmp("ED_BRANCH")
    rsTA("AU_BANK") = rsEmp("ED_BANK")
    rsTA("AU_ACCOUNT") = rsEmp("ED_ACCOUNT")
    rsTA("AU_TRANSITABA") = rsEmp("ED_TRANSITABA")
    rsTA("AU_TRANSITABA2") = rsEmp("ED_TRANSITABA2")
    rsTA("AU_TRANSITABA3") = rsEmp("ED_TRANSITABA3")
    rsTA("AU_AMTDEPOSIT") = rsEmp("ED_AMTDEPOSIT")
    rsTA("AU_PCDEPOSIT") = rsEmp("ED_PCDEPOSIT")
    'BANK 2
    rsTA("AU_DEPOSIT2") = rsEmp("ED_DEPOSIT2")
    rsTA("AU_BRANCH2") = rsEmp("ED_BRANCH2")
    rsTA("AU_BANK2") = rsEmp("ED_BANK2")
    rsTA("AU_ACCOUNT2") = rsEmp("ED_ACCOUNT2")
    rsTA("AU_AMTDEPOSIT2") = rsEmp("ED_AMTDEPOSIT2")
    'BANK3
    rsTA("AU_DEPOSIT3") = rsEmp("ED_DEPOSIT3")
    rsTA("AU_BRANCH3") = rsEmp("ED_BRANCH3")
    rsTA("AU_BANK3") = rsEmp("ED_BANK3")
    rsTA("AU_ACCOUNT3") = rsEmp("ED_ACCOUNT3")
    rsTA("AU_AMTDEPOSIT3") = rsEmp("ED_AMTDEPOSIT3")
    rsTA("AU_PCDEPOSIT3") = rsEmp("ED_PCDEPOSIT3")
    
    rsTA("AU_TD1CODE") = rsEmp("ED_TD1CODE")
    rsTA("AU_TD1DOL") = rsEmp("ED_TD1DOL")
    rsTA("AU_TD3") = rsEmp("ED_TD3")
    rsTA("AU_TD1") = rsEmp("ED_TD1")
    rsTA("AU_DDI") = rsEmp("ED_DDI")
    rsTA("AU_PROVEMP") = rsEmp("ED_PROVEMP")
    rsTA("AU_FedTax") = rsEmp("ED_FedTax")
    rsTA("AU_ExtAmt") = rsEmp("ED_ExtAmt")
    rsTA("AU_ProvForm") = rsEmp("ED_ProvForm")
    rsTA("AU_ProvAmt") = rsEmp("ED_ProvAmt")
    rsTA("AU_ExtraTax") = rsEmp("ED_ExtraTax")
    rsTA("AU_ExtraTaxPC") = rsEmp("ED_ExtraTaxPC")
    rsTA("AU_UIC") = txtUIC
    rsTA("AU_WCB") = rsEmp("ED_WCB")
    rsTA("AU_WCBCODE") = rsEmp("ED_WCBCODE")
    'Employee Status
    rsTA("AU_EMP") = rsEmp("ED_EMP")
    
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbSPCNewEmpNo
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA.Update
    'rsTC.Close
    '------BANK Information End
    
    '------Job and Salary Information
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTC.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        rsTA("AU_JOB") = rsTC("JH_JOB")
        rsTA("AU_DHRS") = rsTC("JH_DHRS")
        rsTA("AU_PHRS") = rsTC("JH_PHRS")
    End If
    rsTC.Close
    rsTC.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        rsTA("AU_SALARY") = rsTC("SH_SALARY")
        rsTA("AU_WHRS") = rsTC("SH_WHRS")
        rsTA("AU_SALCD") = rsTC("SH_SALCD")
        rsTA("AU_SEDATE") = rsTC("SH_EDATE")

        'Ticket #25553 - Pay Period/Company Code change causes Termination and New Hire
        'Use the new Pay Pay/Company Code
        'rsTA("AU_PAYP") = rsTC("SH_PAYP")
        rsTA("AU_PAYP") = glbSPCPPay
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbSPCNewEmpNo
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA.Update
    rsTC.Close
    '------Job and Salary Information END
    
    '------Benefits Begin
    rsTC.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & lblEEID, gdbAdoIhr001, adOpenStatic
    Do While Not rsTC.EOF
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_PTUPL") = rsEmp("ED_PT")
        rsTA("AU_DIVUPL") = rsEmp("ED_DIV")
        rsTA("AU_MTHCCOST") = rsTC("BF_MTHCCOST")
        rsTA("AU_MTHECOST") = rsTC("BF_MTHECOST")
        rsTA("AU_TAXBEN") = rsTC("BF_TAXBEN")
        rsTA("AU_BCODE") = rsTC("BF_BCODE")
        
        rsTA("AU_COVER") = rsTC("BF_COVER")
        rsTA("AU_TCOST") = rsTC("BF_TCOST")
        rsTA("AU_PREMIUM") = rsTC("BF_PREMIUM")
        rsTA("AU_PCE") = rsTC("BF_PCE")
        rsTA("AU_PCC") = rsTC("BF_PCC")
        rsTA("AU_PPAMT") = rsTC("BF_PPAMT")
        rsTA("AU_MAXDOL") = rsTC("BF_MAXDOL")
        rsTA("AU_EDATE") = rsTC("BF_EDATE")
        rsTA("AU_PER") = rsTC("BF_PER")
        rsTA("AU_BAMT") = rsTC("BF_AMT")
        rsTA("AU_UNITCOST") = rsTC("BF_UNITCOST")
        
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbSPCNewEmpNo
        
        rsTA("AU_LDATE") = rsTC("BF_EDATE")
        
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA.Update
        rsTC.MoveNext
    Loop
    rsTC.Close
    '------Benefits End
    
    'ROE Date from HREMP_OTHER table - ER_OTHERDATE1. Only add to new hire (New Employee #) if there is a Date.
    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & lblEEID
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            If rsTA2.State <> 0 Then rsTA2.Close
            rsTA2.Open "SELECT * FROM HRAUDIT2 WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            
            rsTA2.AddNew
            rsTA2("AU_COMPNO") = "001"
            rsTA2("AU_EMPNBR") = glbSPCNewEmpNo
            rsTA2("AU_NEWEMP") = "N"
            
            rsTA2("AU_OTHERDATE1") = rsEmpOther("ER_OTHERDATE1")
            
            rsTA2("AU_LDATE") = Format(Now, "SHORT DATE")
            rsTA2("AU_LUSER") = glbUserID
            rsTA2("AU_LTIME") = Time$
            rsTA2("AU_UPLOAD") = "N"
            rsTA2("AU_TYPE") = "A"
            rsTA2.Update
        End If
    End If
    rsEmpOther.Close
    Set rsEmpOther = Nothing
    
    rsEmp.Close
End Sub

Private Sub cheOUTADDT4_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cboDepositCode_Change()
    If cboDepositCode.ListIndex > -1 Then
        txtDepositCode.Text = Left(cboDepositCode.Text, 1)
    End If
End Sub

Private Sub cboDepositCode_Click()
    If cboDepositCode.ListIndex > -1 Then
        txtDepositCode.Text = Left(cboDepositCode.Text, 1)
    End If
End Sub

Private Sub cboDepositCode2_Change()
    If cboDepositCode2.ListIndex > -1 Then
        txtDepositCode2.Text = Left(cboDepositCode2.Text, 1)
    End If
End Sub

Private Sub cboDepositCode2_Click()
    If cboDepositCode2.ListIndex > -1 Then
        txtDepositCode2.Text = Left(cboDepositCode2.Text, 1)
    End If
End Sub

Private Sub cboDepositCode3_Change()
    If cboDepositCode3.ListIndex > -1 Then
        txtDepositCode3.Text = Left(cboDepositCode3.Text, 1)
    End If
End Sub

Private Sub cboDepositCode3_Click()
    If cboDepositCode3.ListIndex > -1 Then
        txtDepositCode3.Text = Left(cboDepositCode3.Text, 1)
    End If
End Sub

Private Sub cboFedMarry_Change()
    If cboFedMarry.ListIndex > -1 Then
        txtFedMarry.Text = Left(cboFedMarry.Text, 1)
    End If
End Sub

Private Sub cboFedMarry_Click()
    If cboFedMarry.ListIndex > -1 Then
        txtFedMarry.Text = Left(cboFedMarry.Text, 1)
    End If
End Sub

Private Sub cboStateMarry_Change()
    If cboStateMarry.ListIndex > -1 Then
        txtStateMarry.Text = Left(cboStateMarry.Text, 1)
    End If
End Sub

Private Sub cboStateMarry_Click()
    If cboStateMarry.ListIndex > -1 Then
        txtStateMarry.Text = Left(cboStateMarry.Text, 1)
    End If
End Sub

Private Sub chkDirectDeposit_Click()
If chkDirectDeposit.Value = 1 Then
    lblDirectDeposit.Caption = "Y"
Else
    lblDirectDeposit.Caption = "N"
End If

End Sub

Private Sub chkDirectDeposit_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkProvForm_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkProvForm_Click()
If chkProvForm.Value = 1 Then
    lblProvForm.Caption = "Y"
Else
    lblProvForm.Caption = "N"
End If
End Sub

Private Sub chkQTBTORRSP_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkQTBTORRSP_LostFocus()
If chkQTBTORRSP.Value = 1 Then
    lblQTBTORRSP.Caption = "Y"
Else
    lblQTBTORRSP.Caption = "N"
End If
End Sub

Private Sub chkTD1Form_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkTD1Form_Click()

If chkTD1Form.Value = 1 Then
    lblTD1.Caption = "Y"
Else
    lblTD1.Caption = "N"
End If

End Sub

Private Function getProvNo(xProvCode)
Dim rsProv As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = xProvCode
    If Len(xProvCode) > 0 Then
        If Not IsNumeric(xProvCode) Then
            SQLQ = "SELECT * FROM HRPROV WHERE CODE = '" & xProvCode & "' "
            rsProv.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsProv.EOF Then
                If Not IsNull(rsProv("NBR")) Then
                    retVal = rsProv("NBR")
                End If
            End If
        End If
    End If
    getProvNo = retVal
End Function

Private Sub clpProv_LostFocus()
    'Ticket #18417 Samuel, change it back
    'If glbCompSerial = "S/N - 2382W" Then 'Ticket #18082 Samuel
    '    'get prov #
    '    clpProv.Text = getProvNo(clpProv.Text)
    'End If
End Sub


Private Sub cmbCPP_Change()
    If glbInsync Then
        If glbCompSerial = "S/N - 2382W" Then
            txtCPP.Text = Left(cmbCPP.Text, 1)
        ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
            txtCPP.Text = Left(cmbCPP.Text, 1)
        Else
            If cmbCPP.ListIndex = 0 Then
                txtCPP = "0"
            Else
                If glbCompSerial = "S/N - 2292W" And cmbCPP.ListIndex = -1 Then 'County of Elgin
                    txtCPP = ""
                Else
                    txtCPP = "X"
                End If
            End If
        End If
    Else
        txtCPP.Text = cmbCPP.Text
    End If
End Sub

Private Sub cmbCPP_Click()
    If glbInsync Then
        If glbCompSerial = "S/N - 2382W" Then
            txtCPP.Text = Left(cmbCPP.Text, 1)
        Else
            If cmbCPP.ListIndex = 0 Then
                txtCPP = "0"
            Else
                If glbCompSerial = "S/N - 2292W" And cmbCPP.ListIndex = -1 Then 'County of Elgin
                    txtCPP = ""
                Else
                    txtCPP = "X"
                End If
            End If
        End If
    Else
        txtCPP.Text = cmbCPP.Text
    End If
End Sub

Private Sub cmbCPP_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbCPP_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Limit length of data entered to field length
    If Len(cmbCPP.Text) >= 1 And cmbCPP.SelLength = 0 And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
    If glbInsync Then
        If glbCompSerial = "S/N - 2382W" Then
            txtCPP.Text = Left(cmbCPP.Text, 1)
        Else
            If cmbCPP.ListIndex = 0 Then
                txtCPP = "0"
            Else
                If glbCompSerial = "S/N - 2292W" And cmbCPP.ListIndex = -1 Then 'County of Elgin
                    txtCPP = ""
                Else
                    txtCPP = "X"
                End If
            End If
        End If
    Else
        txtCPP.Text = cmbCPP.Text
    End If
End Sub

Private Sub cmbCPP_LostFocus()
    cmbCPP.Text = UCase(cmbCPP.Text)
    'txtCPP.Text = UCase(cmbCPP.Text)
    If glbCompSerial = "S/N - 2382W" Then
        txtCPP.Text = Left(UCase(cmbCPP.Text), 1)
    ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
        txtCPP.Text = Left(UCase(cmbCPP.Text), 1)
    Else
        txtCPP.Text = UCase(cmbCPP.Text)
    End If
End Sub

Private Sub cmbGrossCalc_Click()
'Wellington-Dufferin-Guelph Public Health - Ticket #17129
If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2411W" Then 'St. John's  Ticket #15201
    txtGrossCalc.Text = cmbGrossCalc.Text
End If
End Sub

Private Sub cmbPayFreq_Change()
    txtPAYFREQ.Text = Left(cmbPayFreq.Text, 1)
End Sub

Private Sub cmbPayFreq_Click()
    txtPAYFREQ.Text = Left(cmbPayFreq.Text, 1)
End Sub

Private Sub cmbPayFreq_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbPayFreq_KeyPress(KeyAscii As Integer)
    txtPAYFREQ.Text = Left(cmbPayFreq.Text, 1)
End Sub

Private Sub cmbPenCode_Change()
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Then   'City of Kawartha Lakes & City of Timmins
        txtPension.Text = Left(cmbPenCode.Text, 1)
    Else
        txtPension.Text = cmbPenCode.Text
    End If
End Sub

Private Sub cmbPenCode_Click()
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Then   'City of Kawartha Lakes & City of Timmins
        txtPension.Text = Left(cmbPenCode.Text, 1)
    Else
        txtPension.Text = cmbPenCode.Text
    End If
End Sub

Private Sub cmbPenCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbPenCode_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Limit length of data entered to field length
    If Len(cmbPenCode.Text) >= 1 And cmbPenCode.SelLength = 0 And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Then   'City of Kawartha Lakes & City of Timmins
        txtPension.Text = Left(cmbPenCode.Text, 1)
    Else
        txtPension.Text = cmbPenCode.Text
    End If
End Sub

Private Sub cmbToPayroll_Click()
txtToPayroll.Text = cmbToPayroll.Text
End Sub

Private Sub cmbUIC_Change()
    '2382W - Samuel
    '2383W - Town of Orangeville
    '2425W - Four Villages
    If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2383W" Or glbCompSerial = "S/N - 2425W" Then
        txtUIC.Text = Left(cmbUIC.Text, 1)
    Else
        txtUIC.Text = cmbUIC.Text
    End If
    'Hemu - Vadim
    If glbVadim Then
        'Town of Lasalle - not for them
        If glbCompSerial <> "S/N - 2379W" Then
            If txtUIC.Text = "N" Then
                lblTitle(11).FontBold = False
            Else
                lblTitle(11).FontBold = True
            End If
        End If
    End If
    'Hemu - Vadim
End Sub

Private Sub cmbUIC_Click()
    '2382W - Samuel
    '2383W - Town of Orangeville
    If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2383W" Then
        txtUIC.Text = Left(cmbUIC.Text, 1)
    Else
        txtUIC.Text = cmbUIC.Text
    End If
    'Hemu - Vadim
    If glbVadim Then
        'Town of Lasalle - not for them
        If glbCompSerial <> "S/N - 2379W" Then
            If txtUIC.Text = "N" Then
                lblTitle(11).FontBold = False
            Else
                lblTitle(11).FontBold = True
            End If
        End If
    End If
    'Hemu - Vadim
End Sub

Private Sub cmbUIC_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbUIC_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Limit length of data entered to field length
    If Len(cmbUIC.Text) >= 2 And cmbUIC.SelLength = 0 And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
    If glbInsync Then
        If glbCompSerial = "S/N - 2382W" Then
            txtCPP.Text = Left(cmbCPP.Text, 1)
        Else
            If cmbCPP.ListIndex = 0 Then
                txtCPP = "0"
            Else
                If glbCompSerial = "S/N - 2292W" And cmbCPP.ListIndex = -1 Then 'County of Elgin
                    txtCPP = ""
                Else
                    txtCPP = "X"
                End If
            End If
        End If
    Else
        txtCPP.Text = cmbCPP.Text
    End If
End Sub

Public Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err


rsDATA.CancelUpdate
Call Display_Value

Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEBANK" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err

ODEPCODE = txtDepositCode        'ADDED BY RAUBREY 6/17/97
OBANK = txtBankCode              '
OBRANCH = txtBranchCode          '
OACCOUNT = txtAccount            '
OAMTDEPOSIT = medAmountDeposit   '
OPCDEPOSIT = medPCDeposit        '

ODEPCODE2 = txtDepositCode2      '
OBANK2 = txtBankCode2            '
OBRANCH2 = txtBranchCode2        '
OACCOUNT2 = txtAccount2          '
OAMTDEPOSIT2 = medAmountDeposit2 '
OPCDEPOSIT2 = medPCDeposit2      '

ODEPCODE3 = txtDepositCode3      '
OBANK3 = txtBankCode3            '
OBRANCH3 = txtBranchCode3        '
OACCOUNT3 = txtAccount3          '
OAMTDEPOSIT3 = medAmountDeposit3 '
OPCDEPOSIT3 = medPCDeposit3      '


OGARN = txtGarn            '
OVACPC = medVacPPct
OTD1CODE = txtTD1Code
OTD1DOL = medTD1Amnt
If glbWFC And fraUSA.Visible Then   'And fgetSection(lblEEID) = "GREN" Then
    OGROSCALC = txtFedMarry
    OTD3 = medStateExtra
    OUIC = txtFedExemp
    OCPP = txtStateExemption
    oExtraTax = medFedExtra
    oExtraTaxPC = medFedExtraPC
    oStateExtraTax = medStateExtra
    oStateExtraTaxPC = medStateExtraPC
    OVadim11 = clpVadim11
    OVadim21 = clpVadim21
    oProvCode2 = clpProvE
    OSUPCODE2 = clpCode(0).Text
Else
    OGROSCALC = txtGrossCalc   'laura nov 11, 1997
    OTD3 = medTD3
    OUIC = txtUIC
    OCPP = txtCPP
    oExtraTax = MedExtraTax
    oExtraTaxPC = medExtraTaxPC
End If
OTD1 = lblTD1
OSUPCODE = clpCode(1).Text
ODDI = lblDirectDeposit
oProvEmp = clpProv.Text

OWCB = txtWCB
OWSIBCDE = txtWSIBCde 'Jaddy 11/22/99
If glbWFC And glbCountry = "U.S.A." Then
    OPENSION = txtStatusFalg3
Else
    OPENSION = txtPension
End If
oFedTax = txtFedTax
oExtAmt = txtExtAmt
oProvForm = lblProvForm
oProvAmt = medProvAmt

oProvCode = txtProvCode
OExtrAnn = txtExtrAnn
OQTBTORRSP = lblQTBTORRSP

OOUTADDR = txtOUTAddr
OOUTCITY = txtOUTCity
OOUTPROV = clpOUTProv
OOUTCOUNTRY = comOUTCountry
OOUTPCODE = medOUTPCode
OOUTADDRT4 = chkOUTADDRT4
OTRANSITABA = txtTransitABA
OTRANSITABA2 = txtTransitABA2
OTRANSITABA3 = txtTransitABA3
oVadim1 = clpVadim1
OPenPct = medPenPct
OVadim2 = clpVadim2
oFedAliChd = medCHDSUP
oPAYFREQ = txtPAYFREQ
oHOMEWRKCNT = clpHOME
'Call ST_UPD_MODE(True)

'chkDirectDeposit.SetFocus
Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub
Private Function chkVadimBank()
Dim x
Dim xBNo
Dim SQLQ
On Error GoTo err_This
If glbCompSerial = "S/N - 2362W" Then 'sarnia must be removed if the integration is on
    chkVadimBank = True
    Exit Function
End If
chkVadimBank = False
Dim rsVadimBank As New ADODB.Recordset
For x = 1 To 3
    If x = 1 Then
        xBNo = ""
    Else
        xBNo = x
    End If
    If Me.Controls("txtBankCode" & xBNo) <> "" Or Me.Controls("txtBranchCode" & xBNo) <> "" Then
        SQLQ = "SELECT * FROM BANK_BRANCH "
        SQLQ = SQLQ & " WHERE BANK_CODE ='0" & Me.Controls("txtBankCode" & xBNo) & "'"
        SQLQ = SQLQ & " AND BANK_BRANCH_CODE ='" & Me.Controls("txtBranchCode" & xBNo) & "'"
        rsVadimBank.Open SQLQ, gdbPayroll, adOpenDynamic, adLockReadOnly ' adOpenForwardOnly
        If rsVadimBank.EOF Then
            rsVadimBank.Close
            MsgBox "Invalid Bank Code/Branch #"
            Me.Controls("txtBankCode" & xBNo).SetFocus
            Exit Function
        End If
        rsVadimBank.Close
    End If
Next
chkVadimBank = True

Exit Function
err_This:
If Err.Description = "Invalid object name 'BANK_BRANCH'." Then
    MsgBox "The Banking Information can not be validated because the Vadim table is missing"
End If
chkVadimBank = True
End Function

Private Function chkBank()
Dim Msg, Response, DgDef
Dim xUnion
Dim xPayType
Dim xDept
Dim InvalidUIC, x
Dim xTmpStr

chkBank = False
If glbVadim Then
    If Not chkVadimBank Then Exit Function
End If

'The Walter Fedy Partnership - Ticket #14634
If glbCompSerial = "S/N - 2386W" Then
    If chkDirectDeposit = 1 Then
        If Trim(txtBankCode) = "" Then
           MsgBox "Bank code cannot be blank"
           txtBankCode.SetFocus
           Exit Function
        End If
        If Trim(txtBranchCode) = "" Then
           MsgBox "Branch code cannot be blank"
           txtBranchCode.SetFocus
           Exit Function
        End If
        If Trim(txtAccount) = "" Then
           MsgBox "Account Number cannot be blank"
           txtAccount.SetFocus
           Exit Function
        End If
    End If
End If

'Ticket #18188
'Four Villages Community Health Centre - Ticket #18221
If glbCompSerial = "S/N - 2418W" Or glbCompSerial = "S/N - 2425W" Then
    If Len(clpProv) = 0 Then
        MsgBox lblTitle(9) & " cannot be blank"
        clpProv.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2233W" Then   'Leeds-Grenville F&CS - Ticket #16737
    'If chkDirectDeposit = 1 Then
        If Trim(txtBankCode) = "" Then
           MsgBox "Bank code cannot be blank"
           txtBankCode.SetFocus
           Exit Function
        End If
        If Trim(txtBranchCode) = "" Then
           MsgBox "Branch code cannot be blank"
           txtBranchCode.SetFocus
           Exit Function
        End If
        If Trim(txtAccount) = "" Then
           MsgBox "Account Number cannot be blank"
           txtAccount.SetFocus
           Exit Function
        End If
    'End If
    
    If Len(clpVadim1.Text) = 0 Then
        MsgBox lblVadim1.Caption & " is required field"
        clpVadim1.SetFocus
        Exit Function
    End If
End If

If (glbCompSerial = "S/N - 2409W") Then 'Ticket #30066 Franks - Skylark Children
    If Len(clpCode(1).Text) = 0 Then
        MsgBox lblSupervisor.Caption & " is required field"
        If clpCode(1).Enabled Then clpCode(1).SetFocus
        Exit Function
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox lblSupervisor.Caption & " must be valid"
            If clpCode(1).Enabled Then clpCode(1).SetFocus
            Exit Function
        End If
    End If
End If

If Not glbLinamar Then
    If Not IsNumeric(txtGarn) Then
        MsgBox "Garnishee should be a number"
        txtGarn.SetFocus
        Exit Function
    End If
    'Ticket #18417 Samuel, change it back
    'If glbCompSerial = "S/N - 2382W" Then 'Ticket #18082 Samuel
    '    'get prov #
    '    clpProv.Text = getProvNo(clpProv.Text)
    'End If
    If Len(clpProv.Text) > 0 And clpProv.Caption = "Unassigned" Then
        MsgBox lblTitle(9) & " must be valid"
        clpProv.SetFocus
        Exit Function
    End If
    If glbCountry = "CANADA" And (glbEmpCountry <> "U.S.A.") Then
        If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
            MsgBox lblSupervisor.Caption & " must be valid"
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
End If

If glbWFC Then 'And fgetSection(lblEEID.Caption) = "GREN" Then
    If Len(txtTransitABA.Text) > 0 And Len(txtTransitABA.Text) < 9 Then
        MsgBox lblTitle(1).Caption & " 1 must be 9 characters"
        txtTransitABA.SetFocus
        Exit Function
    End If
    If Len(txtTransitABA2.Text) > 0 And Len(txtTransitABA2.Text) < 9 Then
        MsgBox lblTitle(1).Caption & " 2 must be 9 characters"
        txtTransitABA2.SetFocus
        Exit Function
    End If
    If Len(txtTransitABA3.Text) > 0 And Len(txtTransitABA3.Text) < 9 Then
        MsgBox lblTitle(1).Caption & " 3 must be 9 characters"
        txtTransitABA3.SetFocus
        Exit Function
    End If
    If Len(txtDepositCode.Text) > 0 And Not (txtDepositCode.Text = "V" Or txtDepositCode.Text = "VV" Or txtDepositCode.Text = "W" Or txtDepositCode.Text = "X" Or txtDepositCode.Text = "Y" Or txtDepositCode.Text = "Z" Or txtDepositCode.Text = "T") Then
        MsgBox lblTitle(13).Caption & " 1 must be T, V, VV, W, X, Y or Z" 'Ticket #13574 Frank 08/28/2007, added "T" 'Ticket #29501 Franks 12/16/2016, add "VV"
        txtDepositCode.SetFocus
        Exit Function
    End If
    If Len(txtDepositCode2.Text) > 0 And Not (txtDepositCode2.Text = "V" Or txtDepositCode2.Text = "VV" Or txtDepositCode2.Text = "W" Or txtDepositCode2.Text = "X" Or txtDepositCode2.Text = "Y" Or txtDepositCode2.Text = "Z" Or txtDepositCode.Text = "T") Then
        MsgBox lblTitle(13).Caption & " 2 must be T, V, VV, W, X, Y or Z"
        txtDepositCode2.SetFocus
        Exit Function
    End If
    If Len(txtDepositCode3.Text) > 0 And Not (txtDepositCode3.Text = "V" Or txtDepositCode3.Text = "VV" Or txtDepositCode3.Text = "W" Or txtDepositCode3.Text = "X" Or txtDepositCode3.Text = "Y" Or txtDepositCode3.Text = "Z" Or txtDepositCode.Text = "T") Then
        MsgBox lblTitle(13).Caption & " 3 must be T, V, VV, W, X, Y or Z"
        txtDepositCode3.SetFocus
        Exit Function
    End If
    If glbEmpCountry = "U.S.A." Then
        'Ticket #19266 Franks 12/13/2010
        'move this validation check to Status/Date screen
        'If Len(clpVadim21.Text) < 1 Then
        '    MsgBox lStr("Vadim Field 2 is required field")
        '    clpVadim21.SetFocus
        '    Exit Function
        'ElseIf Len(clpVadim21.Text) > 0 And clpVadim21.Caption = "Unassigned" Then
        '    MsgBox lStr("Vadim Field 2 must be valid")
        '    clpVadim21.SetFocus
        '    Exit Function
        'End If
        If Len(clpProvE.Text) < 1 Then
            MsgBox lblTitle(45) & " is required field"
            clpProvE.SetFocus
            Exit Function
        End If

        'Ticket #22553 Franks 09/18/2012 - begin
        If clpVadim21.Text = "AUY" Or clpVadim21.Text = "AUZ" Then
            If Len(clpHOME.Text) = 0 Then
                'cmbCOMBINATION.ListIndex = 1 '215F
                MsgBox ("Pay Group ") & " is 'AUY' or 'AUZ' then " & "Local Tax Code WI" & " must be '215F'."
                clpHOME.Text = "215F"
                Exit Function
            End If
        Else
            If clpHOME.Text = "215F" Then
                MsgBox ("Pay Group ") & " is not 'AUY' or 'AUZ' then " & "Local Tax Code WI" & " must be blank."
                'cmbCOMBINATION.ListIndex = 0
                clpHOME.Text = ""
                Exit Function
            End If
        End If
        If Len(clpHOME.Text) > 0 Then
            If Len(clpCode(0).Text) = 0 Then
                MsgBox lStr("Supervisor Code") & " is required field if " & "Local Tax Code WI" & " is entered."
                clpCode(0).SetFocus
                Exit Function
            End If
        End If
        'Ticket #22553 Franks 09/18/2012 - end
    End If
End If

If Val(medPCDeposit) > 100 And Len(medPCDeposit) > 0 Then                             'Serbo
      MsgBox "Deposited % can not be more than 100"     '
      medPCDeposit.SetFocus                             '
      Exit Function                                     '
End If                                                  '
If Val(medPCDeposit2) > 100 And Len(medPCDeposit2) > 0 Then                             '
      MsgBox "Deposited % can not be more than 100"     '
      medPCDeposit2.SetFocus                            '
      Exit Function                                     '
End If                                                  '
If Val(medPCDeposit3) > 100 And Len(medPCDeposit3) > 0 Then                             '
      MsgBox "Deposited % can not be more than 100"     '
      medPCDeposit3.SetFocus                            '
      Exit Function                                     '
End If                                                  '

'Frank 12/22/2003,  Surrey Place
If glbCompSerial = "S/N - 2347W" And (Not glbtermopen) Then
    SPCEICode = IIf(IsNull(rsDATA("ED_UIC")), "", rsDATA("ED_UIC"))
    If SPCEICode <> txtUIC.Text Then
        Msg = "EI Code has beed changed" & Chr(10)
        Msg = Msg & "The system will mass change this employee's number. Continue?"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
        Response = MsgBox(Msg, DgDef, "")
        If Response = IDNO Then Exit Function
    End If
End If


If glbPayWeb Then
'    If Val(medTD1Amnt) = 0 Then
'        MsgBox "TD1 Amount is required field."
'        medTD1Amnt.SetFocus
'        Exit Function
'    End If
'    If Val(medProvAmt) = 0 Then
'        MsgBox "Prov. Amount is required field."
'        medProvAmt.SetFocus
'        Exit Function
'    End If
    'If cmbCPP <> "1" And cmbCPP <> "2" And (glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A.") Then
    If cmbCPP <> "1" And cmbCPP <> "2" And cmbCPP <> "Y" And cmbCPP <> "N" And (glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A.") Then
        MsgBox "CPP Code must be 1 or 2 or Y or N"
        cmbCPP.SetFocus
        Exit Function
    End If
    'If cmbUIC <> "1" And cmbUIC <> "2" And (glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A.") Then
    If cmbUIC <> "1" And cmbUIC <> "2" And cmbCPP <> "Y" And cmbCPP <> "N" And (glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A.") Then
        MsgBox "EI Code must be 1 or 2 or Y or N"
        cmbUIC.SetFocus
        Exit Function
    End If
    If txtWCB <> "Y" And txtWCB <> "N" Then
        MsgBox "E.I. Reduce Rate must be Y or N"
        txtWCB.SetFocus
        Exit Function
    ElseIf txtWCB = "Y" Then
        If glbCompSerial <> "S/N - 2381W" Then   'Not The Elliott Community - Ticket #13840
            If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics - Ticket #22716
                If cmbUIC <> "Y" Then
                    MsgBox "EI Code must be Y if E.I. Reduce Rate is equal to Y"
                    cmbUIC.SetFocus
                    Exit Function
                End If
            Else
                If cmbUIC <> "1" Then
                    MsgBox "EI Code must be 1 if E.I. Reduce Rate is equal to Y"
                    cmbUIC.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If txtWSIBCde = "" Then txtWSIBCde = "1"
    If clpProv = "" Then clpProv = "ON"

    If txtBankCode = "" Then
       MsgBox "Bank code is required field."
       txtBankCode.SetFocus
       Exit Function
    End If
    If txtBranchCode = "" Then
       MsgBox "Branch code is required field."
       txtBranchCode.SetFocus
       Exit Function
    End If
    If txtAccount = "" Then
       MsgBox "Account Number is required field."
       txtAccount.SetFocus
       Exit Function
    End If
    If txtBankCode2 <> "" Or txtBranchCode2 <> "" Or txtAccount2 <> "" Or Val(medAmountDeposit2) <> 0 Then
        If txtBankCode2 = "" Then
           MsgBox "Bank code 2 is missing."
           txtBankCode2.SetFocus
           Exit Function
        End If
        If txtBranchCode2 = "" Then
           MsgBox "Branch code 2 is missing."
           txtBranchCode2.SetFocus
           Exit Function
        End If
        If txtAccount2 = "" Then
           MsgBox "Account Number 2 is missing."
           txtAccount2.SetFocus
           Exit Function
        End If
        If Val(medAmountDeposit2) = 0 Then
           MsgBox "Amount Deposit 2 is missing."
           medAmountDeposit2.SetFocus
           Exit Function
        End If
    End If
    
    If txtBankCode3 <> "" Or txtBranchCode3 <> "" Or txtAccount3 <> "" Or Val(medAmountDeposit3) <> 0 Then
        If txtBankCode3 = "" Then
           MsgBox "Bank code 3 is missing."
           txtBankCode3.SetFocus
           Exit Function
        End If
        If txtBranchCode3 = "" Then
           MsgBox "Branch code 3 is missing."
           txtBranchCode3.SetFocus
           Exit Function
        End If
        If txtAccount3 = "" Then
           MsgBox "Account Number 3 is missing."
           txtAccount3.SetFocus
           Exit Function
        End If
        If Val(medAmountDeposit3) = 0 Then
           MsgBox "Amount Deposit 3 is missing."
           medAmountDeposit3.SetFocus
           Exit Function
        End If
    End If
 
End If
If glbInsync Then
    If glbCompSerial <> "S/N - 2411W" And glbCompSerial <> "S/N - 2439W" Then 'Ticket #19392 WDGPHU, OK Tire - Ticket #22128
        If Val(medTD1Amnt) = 0 Then
            MsgBox "TD1 Amount is required field."
            medTD1Amnt.SetFocus
            Exit Function
        End If
    End If
    If glbCompSerial <> "S/N - 2382W" And glbCompSerial <> "S/N - 2411W" And glbCompSerial <> "S/N - 2439W" Then 'Ticket #16478 Samuel and Ticket #19392 WDGPHU, OK Tire - Ticket #22128
        If Val(medProvAmt) = 0 Then
            MsgBox "Prov. Amount is required field."
            medProvAmt.SetFocus
            Exit Function
        End If
    End If
    If glbCompSerial <> "S/N - 2439W" Then  'OK Tire - Ticket #22128
        If Trim(cmbUIC) = "" Then
            MsgBox "EI Code is required field."
            cmbUIC.SetFocus
            Exit Function
        End If
    End If
    
    'Dim InvalidUIC, X
    InvalidUIC = True
    For x = 0 To cmbUIC.ListCount
        If UCase(cmbUIC.List(x)) = UCase(cmbUIC.Text) Then
            InvalidUIC = False
        End If
    Next
    If InvalidUIC Then
        MsgBox "EI Code is invalid."
        cmbUIC.SetFocus
        Exit Function
    End If
    If clpProv = "" Then clpProv = "ON"
    If glbCompSerial = "S/N - 2382W" Then 'Ticket #16478 Samuel
    Else
        If txtBankCode = "" Then
           MsgBox "Bank code is required field."
           txtBankCode.SetFocus
           Exit Function
        End If
        If txtBranchCode = "" Then
           MsgBox "Branch code is required field."
           txtBranchCode.SetFocus
           Exit Function
        End If
        If txtAccount = "" Then
           MsgBox "Account Number is required field."
           txtAccount.SetFocus
           Exit Function
        End If
    End If
    If txtBankCode2 <> "" Or txtBranchCode2 <> "" Or txtAccount2 <> "" Or Val(medAmountDeposit2) <> 0 Then
        If txtBankCode2 = "" Then
           MsgBox "Bank code 2 is missing."
           txtBankCode2.SetFocus
           Exit Function
        End If
        If txtBranchCode2 = "" Then
           MsgBox "Branch code 2 is missing."
           txtBranchCode2.SetFocus
           Exit Function
        End If
        If txtAccount2 = "" Then
           MsgBox "Account Number 2 is missing."
           txtAccount2.SetFocus
           Exit Function
        End If
        'Ticket #18306 Samuel
        'AmountDeposit3 = 0 will stop Deposit Bank 3 at Insync, open it for all Insync users
        'If Val(medAmountDeposit2) = 0 Then
        '   MsgBox "Amount Deposit 2 is missing."
        '   medAmountDeposit2.SetFocus
        '   Exit Function
        'End If
    End If
    
    If txtBankCode3 <> "" Or txtBranchCode3 <> "" Or txtAccount3 <> "" Or Val(medAmountDeposit3) <> 0 Then
        If txtBankCode3 = "" Then
           MsgBox "Bank code 3 is missing."
           txtBankCode3.SetFocus
           Exit Function
        End If
        If txtBranchCode3 = "" Then
           MsgBox "Branch code 3 is missing."
           txtBranchCode3.SetFocus
           Exit Function
        End If
        If txtAccount3 = "" Then
           MsgBox "Account Number 3 is missing."
           txtAccount3.SetFocus
           Exit Function
        End If
        'Ticket #18306 Samuel
        'AmountDeposit3 = 0 will stop Deposit Bank 3 at Insync, open it for all Insync users
        'If Val(medAmountDeposit3) = 0 Then
        '    MsgBox "Amount Deposit 3 is missing."
        '    medAmountDeposit3.SetFocus
        '    Exit Function
        'End If
    End If
 
End If
If glbVadim Then
    If glbCompSerial = "S/N - 2375W" Or _
    txtBankCode <> "" Or txtBranchCode <> "" Or txtAccount <> "" Or Val(medAmountDeposit) <> 0 Or Val(medPCDeposit) <> 0 Then
    '#2375 city of timmins
        If txtBankCode = "" Then
           MsgBox "Bank code is missing."
           txtBankCode.SetFocus
           Exit Function
        End If
        If txtBranchCode = "" Then
           MsgBox "Branch code  is missing."
           txtBranchCode.SetFocus
           Exit Function
        End If
        If txtAccount = "" Then
           MsgBox "Account Number  is missing."
           txtAccount.SetFocus
           Exit Function
        End If
        chkDirectDeposit = 1
    End If
    If txtBankCode2 <> "" Or txtBranchCode2 <> "" Or txtAccount2 <> "" Or Val(medAmountDeposit2) <> 0 Or Val(medPCDeposit3) <> 0 Then
        If txtBankCode2 = "" Then
           MsgBox "Bank code 2 is missing."
           txtBankCode2.SetFocus
           Exit Function
        End If
        If txtBranchCode2 = "" Then
           MsgBox "Branch code 2 is missing."
           txtBranchCode2.SetFocus
           Exit Function
        End If
        If txtAccount2 = "" Then
           MsgBox "Account Number 2 is missing."
           txtAccount2.SetFocus
           Exit Function
        End If
        chkDirectDeposit = 1
    End If
    If txtBankCode3 <> "" Or txtBranchCode3 <> "" Or txtAccount3 <> "" Or Val(medAmountDeposit3) <> 0 Or Val(medPCDeposit3) <> 0 Then
        If txtBankCode3 = "" Then
           MsgBox "Bank code 3 is missing."
           txtBankCode3.SetFocus
           Exit Function
        End If
        If txtBranchCode3 = "" Then
           MsgBox "Branch code 3 is missing."
           txtBranchCode3.SetFocus
           Exit Function
        End If
        If txtAccount3 = "" Then
           MsgBox "Account Number 3 is missing."
           txtAccount3.SetFocus
           Exit Function
        End If
        chkDirectDeposit = 1
    End If
'    If Val(medTD1Amnt) = 0 Then
'        MsgBox "TD1 Amount is required field."
'        medTD1Amnt.SetFocus
'        Exit Function
'    End If
'    If Val(medProvAmt) = 0 Then
'        MsgBox "Prov. Amount is required field."
'        medProvAmt.SetFocus
'        Exit Function
'    End If

    'Ticket #23795 - Town of Lasalle
    If glbCompSerial = "S/N - 2379W" Then
        If cmbUIC <> "01" And cmbUIC <> "02" And cmbUIC <> "00" Then
            MsgBox lblTitle(10).Caption & " must be '00' or '01' or '02'"
            cmbUIC.SetFocus
            Exit Function
        End If
    Else
        If cmbUIC <> "Y" And cmbUIC <> "N" Then
            MsgBox lblTitle(10).Caption & " must be Y or N"
            cmbUIC.SetFocus
            Exit Function
        End If
    End If
    
    If glbCompSerial <> "S/N - 2362W" Then 'sarnia must be removed if the integration is on
        'Ticket #23795 - Not Town of Lasalle
        If glbCompSerial <> "S/N - 2379W" Then
            If cmbUIC = "Y" Then    'Hemu - If Sarnia's serial is removed, please this part of the code for everyone (Vadim)
                If glbCompSerial = "S/N - 2276W" Then
                    If cmbWCB.ListIndex = -1 Then
                        If Trim(cmbWCB.Text) = "" Then
                            MsgBox "E.I. Rate cannot be blank"
                        Else
                            MsgBox "E.I. Rate is invalid"
                        End If
                        cmbWCB.SetFocus
                        Exit Function
                    End If
                ElseIf (glbCompSerial = "S/N - 2363W") Then   'City of Kawartha Lakes
                    If (Left(cmbWCB, 1) <> "1" And Left(cmbWCB, 1) <> "2" And Left(cmbWCB, 1) <> "3") Then
                        MsgBox "E.I. Rate must be 1, 2 or 3"
                        cmbWCB.SetFocus
                        Exit Function
                    End If
                ElseIf cmbWCB <> "1" And cmbWCB <> "2" And cmbWCB <> "3" Then
                    MsgBox "E.I. Rate must be 1, 2 or 3"
                    cmbWCB.SetFocus
                    Exit Function
                End If
            Else
                cmbWCB.Text = ""    'Hemu - if invalid values entered
                cmbWCB.ListIndex = -1
            End If
        End If
'        If cmbWSIBCode.ListIndex = -1 Then
'            MsgBox "WSIB Code is required field"
'            cmbWSIBCode.SetFocus
'            Exit Function
'        End If
        If txtGrossCalc <> "Y" And txtGrossCalc <> "N" Then
            MsgBox "Income Tax Applicable must be Y or N"
            txtGrossCalc.SetFocus
            Exit Function
        End If
        
        If cmbCPP <> "Y" And cmbCPP <> "N" Then
            MsgBox "C.P.P Code must be Y or N"
            cmbCPP.SetFocus
            Exit Function
        End If
    End If
    
    If clpProv = "" Then clpProv = "ON"
    If medVacPPct = "" Then medVacPPct = 0
    
    xUnion = GetEmpData(lblEEID, "ED_ORG")
    xPayType = GetEmpData(lblEEID, "ED_REGION")
    xDept = GetEmpData(lblEEID, "ED_DEPTNO")
    If glbCompSerial = "S/N - 2375W" Then  'City of Timmins
        
        'Ticket #18694
        If IsNumeric(medVacPPct) Then
            If medVacPPct > 0.1 Then
                MsgBox lblTitle(8).Caption & " cannot exceed 10%"
                medVacPPct.SetFocus
                Exit Function
            End If
        End If
    
        If cmbWSIBCode = "" Then
            MsgBox "WSIB Code cannot be blank"
            cmbWSIBCode.SetFocus
            Exit Function
        End If
    
        If xPayType = "H" Then
            If Val(xDept) >= 4402 And Val(xDept) <= 4412 Then
                If Len(clpVadim1.Text) < 1 Then
                    MsgBox lStr("Vadim Field 1 is required field")
                    clpVadim1.SetFocus
                    Exit Function
                ElseIf Len(clpVadim1.Text) > 0 And clpVadim1.Caption = "Unassigned" Then
                    MsgBox lStr("Vadim Field 1 must be valid")
                    clpVadim1.SetFocus
                    Exit Function
                End If
            End If
        End If
        If xUnion = "P" Then
            If Len(clpVadim2.Text) < 1 Then
                MsgBox lStr("Vadim Field 2 is required field")
                clpVadim2.SetFocus
                Exit Function
            ElseIf Len(clpVadim2.Text) > 0 And clpVadim2.Caption = "Unassigned" Then
                MsgBox lStr("Vadim Field 2 must be valid")
                clpVadim2.SetFocus
                Exit Function
            End If
        End If
        
        If IsNumeric(medPenPct) Then
            If medPenPct > 0.14 Then
                MsgBox lblTitle(4).Caption & " cannot exceed 14%"
                medPenPct.SetFocus
                Exit Function
            End If
        End If
        
    End If
End If

If glbCompSerial = "S/N - 2454W" Then   'Showa Canada 'Ticket #24659
    If medVacPPct = "" Then
        MsgBox lblTitle(8).Caption & " is requied. "
        If medVacPPct.Enabled Then medVacPPct.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    If medVacPPct = "" Then
        MsgBox lblTitle(8).Caption & " is requied. "
        If medVacPPct.Enabled Then medVacPPct.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    If cmbUIC = "" Then
        MsgBox "EI Code cannot be blank"
        cmbUIC.SetFocus
        Exit Function
    End If
    
    If cmbWCB = "" Then
        MsgBox "WSIB Code cannot be blank"
        cmbWCB.SetFocus
        Exit Function
    End If
    
    If cmbPenCode = "" Then
        MsgBox "Pension Code cannot be blank"
        cmbPenCode.SetFocus
        Exit Function
    End If
    
    If cmbCPP = "" Then
        MsgBox "CPP Code cannot be blank"
        cmbCPP.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2241W" Then 'Granite Club
    If Len(medVacPPct) = 0 Then medVacPPct = 0
    If Len(clpProv) = 0 Then clpProv = "ON"
    
    InvalidUIC = True
    For x = 0 To cmbUIC.ListCount
        If cmbUIC.List(x) = cmbUIC Then
            InvalidUIC = False
        End If
    Next
    If InvalidUIC Then
        MsgBox "EI Code is invalid."
        cmbUIC.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    If Len(clpProv.Text) < 1 Then
        MsgBox lblTitle(9) & " is required field"
        clpProv.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership
    If Len(clpVadim1.Text) = 0 Then
        MsgBox lblVadim1.Caption & " is required field"
        clpVadim1.SetFocus
        Exit Function
    End If
End If

'Ticket #18770 move this function to Demo screen
'If glbLinamar Then 'Ticket #15216
'    If NewHireForms.count > 0 Then
'        txtToPayroll = "Yes"
'    End If
'End If

If glbCompSerial = "S/N - 2382W" Then 'Ticket #18090 Samuel
    If glbCountry = "CANADA" Then
        If Len(medTD1Amnt.Text) = 0 Then
            MsgBox "TD1 Amount is required field"
            medTD1Amnt.SetFocus
            Exit Function
        End If
        If Len(medProvAmt.Text) = 0 Then
            MsgBox "Prov. Amount is required field"
            medProvAmt.SetFocus
            Exit Function
        End If
        If Len(cmbCPP.Text) = 0 Then
            MsgBox "C.P.P. is required field"
            cmbCPP.SetFocus
            Exit Function
        End If
    End If

    'Ticket #18417 Samuel, change it back
    'clpProv.Text = getProvNo(clpProv.Text)
    
    'Ticket #20600 Franks 09/22/2011 - begin
    If Len(OUIC) > 0 Then
        If Not (OUIC = txtUIC.Text) Then
            Msg = lblTitle(10).Caption & " was changed from '" & OUIC & "' to '" & txtUIC.Text & "' " & Chr(10)
            Msg = Msg & "Are you sure you want to do it?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response = MsgBox(Msg, DgDef, "Confirm Change")
            If Response = IDNO Then Exit Function
        End If
    End If
    If Len(OWCB) > 0 Then
        If Not (OWCB = txtWCB.Text) Then
            Msg = lblTitle(11).Caption & " was changed from '" & OWCB & "' to '" & txtWCB.Text & "' " & Chr(10)
            Msg = Msg & "Are you sure you want to do it?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response = MsgBox(Msg, DgDef, "Confirm Change")
            If Response = IDNO Then Exit Function
        End If
    End If
    If Len(OCPP) > 0 Then
        If Not (OCPP = txtCPP.Text) Then
            Msg = lblTitle(19).Caption & " was changed from '" & OCPP & "' to '" & txtCPP.Text & "' " & Chr(10)
            Msg = Msg & "Are you sure you want to do it?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response = MsgBox(Msg, DgDef, "Confirm Change")
            If Response = IDNO Then Exit Function
        End If
    End If
    If NewHireForms.count = 0 Then 'not for new hire
        'If Len(OWSIBCDE) > 0 Then
            If Not (OWSIBCDE = txtWSIBCde.Text) Then
                Msg = lblTitle(35).Caption & " was changed from '" & OWSIBCDE & "' to '" & txtWSIBCde.Text & "' " & Chr(10)
                Msg = Msg & "Are you sure you want to do it?"
                DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
                Response = MsgBox(Msg, DgDef, "Confirm Change")
                If Response = IDNO Then Exit Function
            End If
        'End If
    End If
    'Ticket #20600 Franks 09/22/2011 - end
End If

'Four Villages Community Health Centre - Ticket #18221
If glbCompSerial = "S/N - 2425W" Then
    If Trim(cmbUIC) = "" Then
        MsgBox "EI Code is required field."
        cmbUIC.SetFocus
        Exit Function
    End If
    
    'If cmbUIC <> "0" And cmbUIC <> "2" And cmbUIC <> "N" And cmbUIC <> "P" And cmbUIC <> "B" And cmbUIC <> "E" And cmbUIC <> "J" And cmbUIC <> "X" Then
    'Ticket #21556 Franks 02/21/2012
    'Ticket #18221
    'If Trim(txtUIC.Text) <> "1" And Trim(txtUIC.Text) <> "2" Then
    If Trim(txtUIC.Text) <> "P" And Trim(txtUIC.Text) <> "N" Then
        MsgBox "Invalid EI Code."
        cmbUIC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(cmbCPP.Text)) = 0 Then
        MsgBox "C.P.P. is required field"
        cmbCPP.SetFocus
        Exit Function
    End If
    
    If cmbCPP <> "0" And cmbCPP <> "X" Then
        MsgBox "Invalid C.P.P."
        cmbCPP.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2453W" Then  'Town of Gander Ticket #24518 Franks 06/04/2015
    If chkTD1Form.Value = 0 Then
        MsgBox "TD1 Form is not checked"
        Exit Function
    End If
    If chkProvForm.Value = 0 Then
        MsgBox "Provincial Form is not checked"
        Exit Function
    End If
    If Len(clpProv.Text) < 1 Then
        MsgBox lblTitle(9).Caption & " is required field"
        clpProv.SetFocus
        Exit Function
    End If
    If Trim(cmbUIC) = "" Then
        MsgBox "EI Code is required field."
        cmbUIC.SetFocus
        Exit Function
    End If
    If Len(Trim(cmbCPP.Text)) = 0 Then
        MsgBox "C.P.P. is required field"
        cmbCPP.SetFocus
        Exit Function
    End If
End If

'Ticket #28786 - Goodmans
If glbCompSerial = "S/N - 2290W" Then
    If Len(clpProv.Text) < 1 Then
        MsgBox lblTitle(9).Caption & " is required field"
        clpProv.SetFocus
        Exit Function
    End If
    If Trim(cmbUIC) = "" Then
        MsgBox "EI Code is required field."
        cmbUIC.SetFocus
        Exit Function
    End If
    If Len(Trim(cmbCPP.Text)) = 0 Then
        MsgBox "C.P.P. is required field"
        cmbCPP.SetFocus
        Exit Function
    End If
End If

'Ticket #20113 - Kerry's Place Autism Services
If glbCompSerial = "S/N - 2433W" Then
    If cmbUIC <> "1" And cmbUIC <> "2" And cmbUIC <> "" Then
        MsgBox "Invalid EI Code."
        cmbUIC.SetFocus
        Exit Function
    End If
    'Ticket #21557
    If cmbCPP <> "Y" And cmbCPP <> "N" Then
        MsgBox "Invalid C.P.P."
        cmbCPP.SetFocus
        Exit Function
    End If
End If

'County of Peterborough - Ticket #28993
If glbCompSerial = "S/N - 2486W" Then
    If cmbUIC <> "1" And cmbUIC <> "2" Then
        MsgBox "Invalid EI Code."
        cmbUIC.SetFocus
        Exit Function
    End If
    
    If cmbCPP <> "1" And cmbCPP <> "2" Then
        MsgBox "Invalid C.P.P."
        cmbCPP.SetFocus
        Exit Function
    End If
    
    'Ticket #30426 Franks 07/26/2017
    'If cmbWCB <> "1" And cmbWCB <> "2" Then
    '    MsgBox "Invalid WCB"
    '    cmbWCB.SetFocus
    '    Exit Function
    'End If
    
    xTmpStr = Left(cmbWSIBCode.Text, 5)
    If xTmpStr <> "590-1" And xTmpStr <> "590-2" And xTmpStr <> "817-1" And xTmpStr <> "817-2" And _
        xTmpStr <> "845-1" And xTmpStr <> "845-2" And xTmpStr <> "COM-1" And xTmpStr <> "COM-2" Then
        MsgBox "Invalid WCB Code"
        cmbWSIBCode.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
    If cmbUIC <> "1" And cmbUIC <> "2" Then
        MsgBox "Invalid EI Code."
        cmbUIC.SetFocus
        Exit Function
    End If
    
    If cmbCPP <> "1" And cmbCPP <> "2" Then
        MsgBox "Invalid C.P.P."
        cmbCPP.SetFocus
        Exit Function
    End If

    If cmbWCB <> "1" And cmbWCB <> "2" Then
        MsgBox "Invalid WCB"
        cmbWCB.SetFocus
        Exit Function
    End If
       
    If cmbWSIBCode <> "" And cmbWSIBCode <> "WCB01" And cmbWSIBCode <> "WCB03" And cmbWSIBCode <> "WCB04" And _
        cmbWSIBCode <> "WCB06" And cmbWSIBCode <> "WCB07" And cmbWSIBCode <> "WCB08" Then
        MsgBox "Invalid WCB Code"
        cmbWSIBCode.SetFocus
        Exit Function
    End If
End If

'If txtDepositCode <> "" Or _
'txtBranchCode <> "" Or _
'txtBankCode <> "" Or _
'txtAccount <> "" Or _
'Val(medAmountDeposit) <> 0 Or _
'Val(medPCDeposit) <> 0 Or _
'txtDepositCode2 <> "" Or _
'txtBranchCode2 <> "" Or _
'txtBankCode2 <> "" Or _
'txtAccount2 <> "" Or _
'Val(medAmountDeposit2) <> 0 Or _
'Val(medPCDeposit2) <> 0 Or _
'txtDepositCode3 <> "" Or _
'txtBranchCode3 <> "" Or _
'txtBankCode3 <> "" Or _
'txtAccount3 <> "" Or _
'Val(medAmountDeposit3) <> 0 Or _
'Val(medPCDeposit3) <> 0 Then
'    chkDirectDeposit = 1
'End If

If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
    InvalidUIC = True
    For x = 0 To cmbUIC.ListCount
        If UCase(cmbUIC.List(x)) = UCase(cmbUIC.Text) Then
            InvalidUIC = False
        End If
    Next
    If InvalidUIC Then
        MsgBox "EI Code is invalid."
        cmbUIC.SetFocus
        Exit Function
    End If

    If cmbCPP <> "0" And cmbCPP <> "X" Then
        MsgBox "Invalid C.P.P."
        cmbCPP.SetFocus
        Exit Function
    End If
End If

'Ticket #19697 - begin
'Franks 02/09/2011
If chkMaxDollar(medAmountDeposit.Text) Then
    If medAmountDeposit.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medAmountDeposit.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(medAmountDeposit2.Text) Then
    If medAmountDeposit2.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medAmountDeposit2.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(medAmountDeposit3.Text) Then
    If medAmountDeposit3.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medAmountDeposit3.SetFocus
    End If
    Exit Function
End If
If Val(medTD3PC) > 100 And Len(medTD3PC) > 0 Then
      MsgBox "Can not be more than 100"     '
      medTD3PC.SetFocus                             '
      Exit Function                                     '
End If

If chkMaxDollar(medTD1Amnt.Text) Then
    If medTD1Amnt.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medTD1Amnt.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(medTD3.Text) Then
    If medTD3.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medTD3.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(medProvAmt.Text) Then
    If medProvAmt.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        medProvAmt.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(MedExtraTax.Text) Then
    If MedExtraTax.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        MedExtraTax.SetFocus
    End If
    Exit Function
End If
If chkMaxDollar(txtExtAmt.Text) Then
    If txtExtAmt.Enabled Then
        MsgBox "Cannot exceed " & glbMaxDollar
        txtExtAmt.SetFocus
    End If
    Exit Function
End If
If Val(medVacPPct) > 100 And Len(medVacPPct) > 0 Then
      MsgBox "Can not be more than 100"     '
      medVacPPct.SetFocus                             '
      Exit Function                                     '
End If
If Val(medExtraTaxPC) > 100 And Len(medExtraTaxPC) > 0 Then
      MsgBox "Can not be more than 100"     '
      medExtraTaxPC.SetFocus                             '
      Exit Function                                     '
End If
If Len(medPenPct.Text) > 0 Then
    If Not IsNumeric(medPenPct.Text) Then
      MsgBox "Invalid Percentage"     '
      medPenPct.SetFocus                             '
      Exit Function
    Else
        If Val(medPenPct) > 100 And Len(medPenPct) > 0 Then
            MsgBox "Can not be more than 100"     '
            medPenPct.SetFocus                             '
            Exit Function
        End If
    End If
End If
'Ticket #19697 - end

'Ticket #20931 - Town of Aurora
If glbCompSerial = "S/N - 2378W" Then
    If cmbWSIBCode.ListIndex = -1 Then
        MsgBox lblTitle(35).Caption & " is required field"
        cmbWSIBCode.SetFocus
        Exit Function
    End If
End If

'Ticket #25670 - District of Muskoka
If glbCompSerial = "S/N - 2373W" Then
    If cmbWSIBCode.ListIndex = -1 Then
        MsgBox lblTitle(35).Caption & " is required field"
        cmbWSIBCode.SetFocus
        Exit Function
    End If
End If

chkBank = True
End Function

Public Sub cmdOK_Click()
Dim DtTm As Variant, rc As Integer
Dim x

DtTm = Now
rsDATA.Requery
chkDirectDeposit.SetFocus

On Error GoTo Add_Err

'St. John's Rehab Hospital - Ticket #14685
If glbCompSerial = "S/N - 2394W" Then
    If Trim(txtBankCode.Text) <> "" Or Trim(txtBranchCode.Text) <> "" Or Trim(txtAccount.Text) <> "" Then
        txtDepositCode.Text = "1"
    Else
        txtDepositCode.Text = ""
    End If
    If Trim(txtBankCode2.Text) <> "" Or Trim(txtBranchCode2.Text) <> "" Or Trim(txtAccount2.Text) <> "" Then
        txtDepositCode2.Text = "2"
    Else
        txtDepositCode2.Text = ""
    End If
    If Trim(txtBankCode3.Text) <> "" Or Trim(txtBranchCode3.Text) <> "" Or Trim(txtAccount3.Text) <> "" Then
        txtDepositCode3.Text = "3"
    Else
        txtDepositCode3.Text = ""
    End If
End If

If Not chkBank Then Exit Sub

rsDATA.Requery
txtOUTCountry = comOUTCountry

'Frank 12/22/2003,  Surrey Place
If glbCompSerial = "S/N - 2347W" And (Not glbtermopen) Then
    glbSPCTermDate = ""
    glbSPCTermReason = ""
    glbSPCPPay = ""     'Ticket #25553 - Pay Period/Company Code change causes Termination and New Hire
    SPCEICode = IIf(IsNull(rsDATA("ED_UIC")), "", rsDATA("ED_UIC"))
    'Screen.MousePointer = Default
    If SPCEICode <> txtUIC.Text Then
        If Len(SPCEICode) > 0 And Len(txtUIC) Then
            frmSTermPara.Show 1
        End If
    End If
    'Screen.MousePointer = HOURGLASS
End If

If glbInsync Then
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18417
    Else
        'Ticket #18437 May 7, 2010 Frank
        'If cmbWCB.ListIndex = 0 Then
        '    txtWCB = "0"
        'Else
        '    txtWCB = "M"
        'End If
    End If
End If
If glbVadim Then
    'Town of Lasalle - not for them
    If glbCompSerial <> "S/N - 2379W" Then
        txtWCB = Left(cmbWCB, 1)
    End If
    If cmbWSIBCode.Visible Then
        If cmbWSIBCode.ListIndex > -1 Then
            txtWSIBCde = cmbWSIBCode.ItemData(cmbWSIBCode.ListIndex)
        End If
    End If
End If
If Not glbtermopen Then
    If Not AUDITBANK() Then MsgBox "ERROR : AUDIT FILE"
ElseIf glbVadim And glbtermopen Then
    Call Pass_TermBank_Changes_Vadim
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)

'removed by jaddy because this cause problems
'If Not glbLinamar Then
'    lblQTBTORRSP = clpProv
'    txtExtrAnn = medVacPPct
'Else
'    clpProv = lblQTBTORRSP
'    medVacPPct = txtExtrAnn
'End If

If glbWFC Then 'And fgetSection(lblEEID) = "GREN" Then
    If Not IsNull(rsDATA("ED_TD3PC")) And Len(medStateExtraPC) > 0 Then rsDATA("ED_TD3PC") = Int(Val(medStateExtraPC * 100))
    If Not IsNull(rsDATA("ED_EXTRATAXPC")) And Len(medFedExtraPC) > 0 Then rsDATA("ED_EXTRATAXPC") = Int(Val(medFedExtraPC * 100))
Else
    If IsNull(rsDATA("ED_TD3PC")) = False And IsNumeric(medTD3PC) = True Then rsDATA("ED_TD3PC") = Int(Val(medTD3PC * 100))
    If IsNull(rsDATA("ED_EXTRATAXPC")) = False And IsNumeric(medExtraTaxPC) = True Then rsDATA("ED_EXTRATAXPC") = Int(Val(medExtraTaxPC * 100))
End If
If Not IsNull(rsDATA("ED_TD1DOL")) Then rsDATA("ED_TD1DOL") = Int(Val(medTD1Amnt))

'City of Timmins - For RPP # (Vadim)
If glbCompSerial = "S/N - 2375W" Then
    If txtPension = "1" Then
        rsDATA("ED_NORMALR") = DateAdd("yyyy", 65, CVDate(rsDATA("ED_DOB")))
    ElseIf txtPension = "2" Then
        rsDATA("ED_NORMALR") = DateAdd("yyyy", 60, CVDate(rsDATA("ED_DOB")))
    End If
End If

Call Set_Control("U", Me, rsDATA)
rsDATA.Update
Data1.Refresh
Call Display_Value

If glbMediPay Then 'Ticket #14752
    Call Bank_Integration(glbLEE_ID)
End If

If glbGP Then 'Ticket #17641 Frank 12/04/09
    Call Employee_Master_Integration(glbLEE_ID)
End If

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'Frank 12/22/2003 for Surrey Place
'If EI Code changed, enter a new Employee # and do the Employee # Mass change
If glbCompSerial = "S/N - 2347W" Then
    If Len(glbSPCTermReason) > 0 Then
        Call modEmpNoUpdate
        glbLEE_ID = glbSPCNewEmpNo
        
        'Ticket #25553 - Pay Period/Company Code change causes Termination and New Hire
        'Update Current Salary record with the new Pay Period/Comapny Code
        Call SPCUpdatePayPeriod(glbLEE_ID, glbSPCPPay)
        
        Unload frmEBANK
        Call UnloadFrms
    End If
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    Call FamilyDayEmpSync(glbLEE_ID)
End If

Call NextForm
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub ComPenCode()

If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    cmbPenCode.AddItem "5"
    cmbPenCode.AddItem "6"
    cmbPenCode.AddItem "E"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2357W" And glbEmpCountry = "U.S.A." Then   'I.T. Xchange
    cmbPenCode.Clear
    cmbPenCode.AddItem "S"
    cmbPenCode.AddItem "M"
    cmbPenCode.AddItem "H"
    cmbPenCode.AddItem "J"
    cmbPenCode.AddItem "X"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2347W" Then   'Surrey Place
    cmbPenCode.AddItem "7"
    cmbPenCode.AddItem "8"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
    cmbPenCode.Clear
    cmbPenCode.Width = 2745
    cmbPenCode.AddItem "1 - Retirement Age 65"
    cmbPenCode.AddItem "2 - Retirement Age 60"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2375W" Then 'City of Timmins
    cmbPenCode.Clear
    cmbPenCode.Width = 2745
    cmbPenCode.AddItem "1 - Retirement Age 65"
    cmbPenCode.AddItem "2 - Retirement Age 60"
ElseIf glbCompSerial = "S/N - 2385W" Then   'Conservation Halton 'Ticket #13063
    cmbPenCode.Clear
    cmbPenCode.AddItem "5"
    cmbPenCode.AddItem "E"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2447W" Then   'Town of Greater Napanee 'Ticket #22905
    cmbPenCode.Clear
    cmbPenCode.AddItem "Y"
    cmbPenCode.AddItem "N"
    cmbPenCode.AddItem ""
ElseIf glbCompSerial = "S/N - 2379W" Then   'Town of Lasalle
    cmbPenCode.AddItem "1"
    cmbPenCode.AddItem "2"
    cmbPenCode.AddItem "E"
    cmbPenCode.AddItem ""
End If
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Banking Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Banking Information Report"

Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgbank.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
    End If
    xReport = glbIHRREPORTS & "rgbank2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub
Public Sub cmdView_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Banking Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Banking Information Report"

Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgbank.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
    End If
    xReport = glbIHRREPORTS & "rgbank2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub
'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmbUIC_LostFocus()
    cmbUIC.Text = UCase(cmbUIC.Text)
    If glbCompSerial = "S/N - 2382W" Then
        txtUIC.Text = Left(UCase(cmbUIC.Text), 1)
    Else
        txtUIC.Text = UCase(cmbUIC.Text)
    End If
End Sub

Private Sub cmbWCB_Click()
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18417
    txtWCB.Text = Left(cmbWCB.Text, 1)
ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
    txtWCB.Text = Left(cmbWCB.Text, 1)
ElseIf glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    txtWCB.Text = cmbWCB.Text
Else
    txtWCB.Text = cmbWCB.Text
End If
End Sub

Private Sub cmbWCB_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbWCB_KeyPress(KeyAscii As Integer)
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18417
Else
'If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    If Len(cmbWCB.Text) >= 2 And cmbWCB.SelLength = 0 And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
    txtWCB.Text = cmbWCB.Text
'End If
End If
End Sub

Private Sub cmbWSIBCode_Change()
    'Ticket #18417 Samuel, County of Peterborough - Ticket #28993, City of Kenora - Ticket #29054
    If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2487W" Then
    Else
        If cmbWSIBCode.ListIndex > -1 And doOnce = False Then
            If glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
                txtWSIBCde.Text = Left(cmbWSIBCode.Text, 5)
            ElseIf glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
                txtWSIBCde.Text = cmbWSIBCode.Text
            Else
                txtWSIBCde = cmbWSIBCode.ItemData(cmbWSIBCode.ListIndex)
            End If
            doOnce = True
        ElseIf cmbWSIBCode.ListIndex = -1 Then
            txtWSIBCde = ""
        End If
    End If
End Sub

Private Sub cmbWSIBCode_Click()
    If cmbWSIBCode.ListIndex > -1 Then
        If glbCompSerial = "S/N - 2382W" Then 'Ticket #18417 Samuel
            txtWSIBCde = Left(cmbWSIBCode.Text, 1)
        ElseIf glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
            'txtWSIBCde.Text = cmbWSIBCode.Text
            txtWSIBCde.Text = Left(cmbWSIBCode.Text, 5)
        ElseIf glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
            txtWSIBCde.Text = cmbWSIBCode.Text
        Else
            txtWSIBCde = cmbWSIBCode.ItemData(cmbWSIBCode.ListIndex)
        End If
    End If

End Sub
Private Sub cmbWSIBCode_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub comOUTCountry_Click()
If Not glbLinamar Then Exit Sub
If comOUTCountry = "CANADA" Then                   'laura Oct 28, 1997
    lblTitle(30) = "Province"                    '
    clpOUTProv.Tag = "31-Province - Code"          '
    lblTitle(31) = "Postal Code"                 '
    medOUTPCode.Tag = "01-Postal Code"             '
    medOUTPCode.MaxLength = 7
    medOUTPCode.Mask = "?#? #?#"

ElseIf comOUTCountry = "BAHAMAS" Then              '
    medOUTPCode.MaxLength = 8
    medOUTPCode.Mask = "AAAAAAAA"
    lblTitle(30) = "Island"                      '
    clpOUTProv.Tag = "30-Island - Code"            '
    lblTitle(31) = "Postal Code"                 '
    medOUTPCode.Tag = "01-Postal Code"             '
ElseIf comOUTCountry = "U.S.A." Or comOUTCountry = "MEXICO" Then
    lblTitle(30) = "State"                   '
    clpOUTProv.Tag = "31-State - Code"         '
    lblTitle(31) = "Zip Code"                '
    medOUTPCode.Tag = "01-Zip Code"            '
    medOUTPCode.MaxLength = 10
    medOUTPCode.Mask = "AAAAA-AAAA"
Else
    lblTitle(30) = "Province"                '
    clpOUTProv.Tag = "31-Province - Code"      '
    lblTitle(31) = "Postal Code"             '
    medOUTPCode.Tag = "01-Postal Code"         '
    medOUTPCode.Mask = "&&&&&&&&&&&&&&&"
    medOUTPCode.MaxLength = 15 ' 10
End If                                          '

End Sub
Private Sub comOUTCountry_GotFocus() 'RAUBREY 6/16/97
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comOUTCountry_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRSTATS", "SELECT")

End Sub


Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

'If glbtermopen Then  'Lucy July 4, 2000
'    SQLQ = "Select " & FldList & " from Term_HREMP"  'Lucy July 11, 2000
'    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'Else
'    SQLQ = "Select " & FldList & " from HREMP "    'Lucy July 11, 2000
'    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
'End If
'data1.RecordSource = SQLQ
'data1.Refresh



If glbtermopen Then
    SQLQ = "Select " & FldList & " from Term_HREMP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select " & FldList & " from HREMP "
    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Call Display_Value

EERetrieve = True

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    lblTitle(9).FontBold = True
ElseIf glbCompSerial = "S/N - 2357W" And glbEmpCountry <> "CANADA" Then    'I.T. Xchange
    lblTitle(9).FontBold = False
End If

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack

Exit Function

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click

If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership
    If (NewHireForms.count > 0) Or fglbNew Then     'for New Hire only - Ticket #14634
        chkDirectDeposit = 1
    End If
End If

End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEBANK"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found


Screen.MousePointer = HOURGLASS
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbLinamar Then
    lblQTBTORRSP.DataField = "ED_QTBTORRSP"
    txtExtrAnn.DataField = "ED_EXTRANN"
    'Ticket #15216 - begin
    clpVadim2.DataField = ""
    txtToPayroll.DataField = "ED_VADIM2"
    cmbToPayroll.AddItem "Yes"
    cmbToPayroll.AddItem "No"
    cmbToPayroll.ListIndex = 0
    lblTitle(41).Visible = True
    cmbToPayroll.Visible = True
    'Ticket #15216 - end
    
    'Ticket #20188 - Begin - Increase the Bank Code for non Canadian and US employees to 5
    If glbEmpCountry <> "CANADA" And glbEmpCountry <> "U.S.A." Then
        txtBankCode.MaxLength = 5
        txtBankCode2.MaxLength = 5
        txtBankCode3.MaxLength = 5
    End If
End If

Call setCaption(lblSupervisor)
'Call setCaption(lblVadim1)
'Call setCaption(lblVadim2)
lblVadim1.Caption = lStr("Vadim Field 1")
lblVadim2.Caption = lStr("Vadim Field 2")
If glbWFC Then 'Ticket #16392
    lblVadim11.Caption = lStr("Vadim Field 1")
    lblVadim21.Caption = lStr("Vadim Field 2")
    lblTitle(16).Caption = "Full" 'Bank 1
    lblTitle(17).Caption = "Partial 1" 'Bank 2
    lblTitle(18).Caption = "Partial 2" 'Bank 3
    'Ticket #19306 - begin
    'move Vadim 1 and Vadim 2 to Status/Dates screen for NGS Sub group
    'So disable these two fields here, user can't change them
    lblVadim11.Enabled = False
    lblVadim21.Enabled = False
    clpVadim11.Enabled = False
    clpVadim21.Enabled = False
    'Ticket #19306 - end
    
    'Ticket #22553 Franks 09/17/2012 - being
    lblTitle(46).Caption = lStr("Supervisor Code")
    lblTitle(47).Caption = "Local Tax Code WI" ' lStr("Combination") ' "Local Tax Code WI"
    'Ticket #22553 Franks 09/17/2012 - end
    
    'Ticket #25969 - Franks 09/09/2014 - "   On Bank 1, hide Amount Deposit and % Deposit boxes.
    medAmountDeposit.Visible = False
    medPCDeposit.Visible = False
End If

If glbInsync Then
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18417
        cmbWCB.Clear
        cmbWCB.Width = 2745
        cmbWCB.AddItem "0 - Default ET Code"
        cmbWCB.AddItem "A - Tax on T4A / Releve1"
        cmbWCB.AddItem "B - Combination of options A and P"
        cmbWCB.AddItem "C - TR is a % Tax Credit"
        cmbWCB.AddItem "F - Calc Fed Tax/Exempt TONI"
        cmbWCB.AddItem "H - Que Emp living in Ontario"
        cmbWCB.AddItem "I - Status Indian"
        cmbWCB.AddItem "M - Exempt Fed & Prov Tax"
        cmbWCB.AddItem "N - NWT & Nunavut exempt"
        cmbWCB.AddItem "O - Default ET Code"
        cmbWCB.AddItem "P - TR becomes a %"
        cmbWCB.AddItem "Q - Exempt Fed Tax/Pays Que Tax"
        cmbWCB.AddItem "T - TA becomes a %"
        cmbWCB.AddItem "X - Exempt Fed Tax Only"
        cmbWCB.ListIndex = -1
    ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
        cmbWCB.Clear
        cmbWCB.Width = 1500
        cmbWCB.AddItem "   - Normal"
        cmbWCB.AddItem "X - Exempt"
        cmbWCB.ListIndex = -1
    Else
        cmbWCB.AddItem "0"
        cmbWCB.AddItem "M"
        cmbWCB.AddItem "O"
        cmbWCB.AddItem "P"
        cmbWCB.AddItem "X"
        cmbWCB.ListIndex = 0
    End If
ElseIf glbVadim Then
    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
        cmbWCB.Width = 2745
        'cmbWCB.AddItem "1 - City Special"
        'cmbWCB.AddItem "2 - City Normal"
        ''cmbWCB.AddItem "3 - Humane Society"
        'cmbWCB.AddItem "3 - Library Special" 'Ticket #22926 Franks 12/05/2012
        'cmbWCB.AddItem "4 - Library"
        'cmbWCB.AddItem "5 - Transit Special"
        
        'Ticket #28401 - Changed to...
        cmbWCB.AddItem "1 - NU/CUPE Permanent"
        cmbWCB.AddItem "2 - PT/TEMPS/Casuals/Seasonals"
        cmbWCB.AddItem "3 - Library FT"
        cmbWCB.AddItem "4 - Library PT"
        cmbWCB.AddItem "5 - Transit Permanent"
        cmbWCB.AddItem "6 - FIRE"
        
    ElseIf glbCompSerial = "S/N - 2363W" Then   'Kawartha Lakes
        cmbWCB.Width = 2745
        cmbWCB.AddItem "1. Unreduced (No Sick Pay Plan)"
        cmbWCB.AddItem "2. FT Salaried with Sick Pay Plan"
        cmbWCB.AddItem "3. FT Union with Sick Pay Plan"
    ElseIf glbCompSerial = "S/N - 2373W" Then 'Ticket #24565- District Municipality of South Muskoka
        cmbWCB.AddItem "1"
        cmbWCB.AddItem "2"
        'Ticket #28927
        cmbWCB.AddItem "3"
    Else
        cmbWCB.AddItem "1"
        cmbWCB.AddItem "2"
        cmbWCB.AddItem "3"
    End If
Else
    If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
        cmbWCB.AddItem "1"
        cmbWCB.AddItem "2"
        cmbWCB.AddItem "3"
        cmbWCB.AddItem "4"
        cmbWCB.AddItem "6"
        cmbWCB.AddItem ""
        cmbWCB.Visible = True
    End If
    If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics - Ticket #22716
        cmbWCB.AddItem "Y"
        cmbWCB.AddItem "N"
        cmbWCB.Visible = True
    End If
    
    'Ticket #30426 Franks 07/26/2017
    'If glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
    '    cmbWCB.AddItem "1"
    '    cmbWCB.AddItem "2"
    '    cmbWCB.Visible = True
    'End If
    If glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
        cmbWCB.AddItem "1"
        cmbWCB.AddItem "2"
        cmbWCB.Visible = True
    End If
End If

If glbVadim Then
    'WSIB codes must start with a 1, OR an array hast to be created to equate the
    'first code with Index 0. Reccommend using the Itemdata property or
    'in VB.net creating an index property on the list item.
    
    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
        cmbWSIBCode.AddItem "1 - Library"
        cmbWSIBCode.ItemData(0) = 1
        'cmbWSIBCode.AddItem "2 - Museum"    'Ticket #29486 - They do not want this anymore
        'cmbWSIBCode.ItemData(1) = 2
        'cmbWSIBCode.AddItem "3 - Humane Society"    'Ticket #29486 - They do not want this anymore
        'cmbWSIBCode.ItemData(2) = 3
        lblTitle(21).Caption = "Sick Accrual Hours"
        medCHDSUP.Tag = "00-Sick Accrual Hours"
        
        'For Pay Vac. Every Period
        cmbPayFreq.AddItem "Yes"
        cmbPayFreq.AddItem "No"
        cmbPayFreq.Visible = True
        txtPAYFREQ.Visible = False
        
    ElseIf glbCompSerial = "S/N - 2378W" Then 'Aurora
        'added by Bryan 25/Oct/05 Ticket#9607
        'Changed by Bryan 22/Nov/05 Ticket#9832
        'cmbWSIBCode.AddItem "1 - Historical Society"
        cmbWSIBCode.AddItem "1 - Town"
        cmbWSIBCode.ItemData(0) = 1
        cmbWSIBCode.AddItem "2 - Library"
        cmbWSIBCode.ItemData(1) = 2
        cmbWSIBCode.AddItem "3 - No WSIB"
        cmbWSIBCode.ItemData(2) = 3
    ElseIf glbCompSerial = "S/N - 2375W" Then 'City of Timmins
        cmbWSIBCode.AddItem "1 - Misc General"
        cmbWSIBCode.ItemData(0) = 1
        cmbWSIBCode.AddItem "2 - Golden Manor"
        cmbWSIBCode.ItemData(1) = 2
        cmbWSIBCode.AddItem "3 - Library"
        cmbWSIBCode.ItemData(2) = 3
        cmbWSIBCode.AddItem "4 - Police"
        cmbWSIBCode.ItemData(3) = 4
        cmbWSIBCode.AddItem "5 - Transit"
        cmbWSIBCode.ItemData(4) = 5
        'cmbWSIBCode.AddItem ""
        'cmbWSIBCode.ItemData(5) = 0
        cmbWSIBCode.AddItem "7 - Museum"
        cmbWSIBCode.ItemData(5) = 7
        cmbWSIBCode.AddItem "8 - Water Filtration"
        cmbWSIBCode.ItemData(6) = 8
        cmbWSIBCode.AddItem "9 - MRCA"
        cmbWSIBCode.ItemData(7) = 9
        cmbWSIBCode.AddItem "10 - Airport"
        cmbWSIBCode.ItemData(8) = 10
        cmbWSIBCode.AddItem "11 - Exempt"
        cmbWSIBCode.ItemData(9) = 11
    ElseIf glbCompSerial = "S/N - 2373W" Then 'District of Muskoka - Ticket #25670
        cmbWSIBCode.AddItem "1 - General Municipal/Regional OPS"
        cmbWSIBCode.ItemData(0) = 1
        'Ticket #28927
        cmbWSIBCode.AddItem "2 - Paramedics"
        cmbWSIBCode.ItemData(1) = 2
    ElseIf glbCompSerial = "S/N - 2458W" Then 'Ticket #25469 - City of Campbell River
        cmbWSIBCode.AddItem "Y - Yes"
        cmbWSIBCode.ItemData(0) = 1
        cmbWSIBCode.AddItem "N - No"
        cmbWSIBCode.ItemData(1) = 2
    ElseIf glbCompSerial = "S/N - 2447W" Then 'Ticket #25412 - Town of Greater Napanee
        cmbWSIBCode.AddItem "1 - Town"
        cmbWSIBCode.ItemData(0) = 1
        cmbWSIBCode.AddItem "2 - Utilities"
        cmbWSIBCode.ItemData(1) = 2
    Else
        cmbWSIBCode.AddItem "1 - General Municipal/Regional OPS"
        cmbWSIBCode.ItemData(0) = 1
        cmbWSIBCode.AddItem "2 - Public Health Clinics"
        cmbWSIBCode.ItemData(1) = 2
        cmbWSIBCode.AddItem "3 - Child Day Care"
        cmbWSIBCode.ItemData(2) = 3
        cmbWSIBCode.AddItem "4 - Land Ambulance"
        cmbWSIBCode.ItemData(3) = 4
        cmbWSIBCode.AddItem "5 - Operators of Apartments"
        cmbWSIBCode.ItemData(4) = 5
        cmbWSIBCode.AddItem "6 - Libraries & Museums"
        cmbWSIBCode.ItemData(5) = 6
        cmbWSIBCode.AddItem "7 - Nursing Home Operations"
        cmbWSIBCode.ItemData(6) = 7
        'Ticket #20151
        cmbWSIBCode.AddItem "8 - Tourism Sarnia Lambton"
        cmbWSIBCode.ItemData(7) = 8
    End If
Else
    If glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
        'cmbWSIBCode.Width = ""
        cmbWSIBCode.AddItem "590-1 AMBULANCE OPERATION-PT"
        'cmbWSIBCode.ItemData(0) = "1"
        cmbWSIBCode.AddItem "590-2 AMBULANCE OPERATION-FT"
        'cmbWSIBCode.ItemData(1) = "2"
        cmbWSIBCode.AddItem "817-1 MUSEUM & ARCHIVES-PT"
        'cmbWSIBCode.ItemData(2) = "3"
        cmbWSIBCode.AddItem "817-2 MUSEUM & ARCHIVES-FT"
        'cmbWSIBCode.ItemData(3) = "4"
        cmbWSIBCode.AddItem "845-1 GENERAL NUNICIPAL/REG,OPS-PT"
        'cmbWSIBCode.ItemData(4) = "5"
        cmbWSIBCode.AddItem "845-2 GENERAL NUNICIPAL/REG,OPS-FT"
        'cmbWSIBCode.ItemData(5) = "6"
        cmbWSIBCode.AddItem "COM-1 COMMON EARNINGS - PT"
        'cmbWSIBCode.ItemData(6) = "7"
        cmbWSIBCode.AddItem "COM-2 COMMON EARNINGS - FT"
        'cmbWSIBCode.ItemData(7) = "8"
    End If
    If glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
        cmbWSIBCode.AddItem ""
        'cmbWSIBCode.ItemData(0) = "1"
        cmbWSIBCode.AddItem "WCB01"
        'cmbWSIBCode.ItemData(1) = "2"
        cmbWSIBCode.AddItem "WCB03"
        'cmbWSIBCode.ItemData(2) = "3"
        cmbWSIBCode.AddItem "WCB04"
        'cmbWSIBCode.ItemData(3) = "4"
        cmbWSIBCode.AddItem "WCB06"
        'cmbWSIBCode.ItemData(4) = "5"
        cmbWSIBCode.AddItem "WCB07"
        'cmbWSIBCode.ItemData(5) = "6"
        cmbWSIBCode.AddItem "WCB08"
        'cmbWSIBCode.ItemData(6) = "7"
    End If
End If
If glbInsync Then
    If glbCompSerial = "S/N - 2382W" Then 'Ticket #16235 Samuel
        cmbCPP.Width = 2500
        cmbCPP.AddItem "0 - Default EP Code"
        cmbCPP.AddItem "E - No CPP/QPP Exemption"
        cmbCPP.AddItem "M - QPP even if PC not 4"
        cmbCPP.AddItem "X - Exempt"
    ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
        cmbCPP.Width = 1500
        cmbCPP.Clear
        cmbCPP.AddItem "   - Normal"
        cmbCPP.AddItem "X - Exempt"
        cmbCPP.ListIndex = -1
    Else
        If glbCompSerial = "S/N - 2295W" Then
            cmbCPP.AddItem "O" '"0" Ticket# 7118
        Else
            cmbCPP.AddItem "0"
        End If
        cmbCPP.AddItem "X"
    End If
    
    If glbCompSerial = "S/N - 2292W" Then   'County of Elgin
        cmbCPP.ListIndex = -1
    Else
        cmbCPP.ListIndex = 0
    End If
Else
    If glbVadim Then    'Hemu
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        If glbCompSerial = "S/N - 2378W" Then 'Town of Aurora - Ticket #20931
            cmbCPP.ListIndex = 0
        End If
    ElseIf glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2386W" Then   'The Walter Fedy Partnership
        cmbCPP.AddItem "0"
        cmbCPP.AddItem "X"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2385W" Then   ' Conservation Halton 'Ticket #13063
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2381W" Then   ' The Elliott Community 'Ticket #13603
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2394W" Then 'St. John's  Ticket #15201
        cmbCPP.AddItem ""
        cmbCPP.AddItem "N"
        cmbCPP.AddItem "S"
        cmbCPP.AddItem "Y"
    ElseIf glbCompSerial = "S/N - 2414W" Then   ' Merrickville DCHSC 'Ticket #17210
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2425W" Then   ''Four Villages Community Health Centre - Ticket #18221
        cmbCPP.AddItem "0"
        cmbCPP.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Autism Services  Ticket #20113
        'cmbCPP.AddItem "1"
        'cmbCPP.AddItem ""
        'Ticket #21504
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2396W" Then 'Oshawa CHC Ticket #20598 Franks 06/29/2012
        cmbCPP.AddItem "1"
        cmbCPP.AddItem "2"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics - Ticket #22716
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2447W" Then   'Town of Greater Napanee - Ticket #22905
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        cmbCPP.AddItem ""
    ElseIf glbCompSerial = "S/N - 2457W" Then   'McLeod Law 'Ticket #24863
        'cmbCPP.AddItem "Y"
        'cmbCPP.AddItem "N"
        'Ticket #24864 Franks 06/10/2014
        cmbCPP.Clear
        cmbCPP.AddItem "0"
        cmbCPP.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2453W" Then   'Town of Gander 'Ticket #24518 Franks 04/15/2014
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2454W" Then   'Showa Canada 'Ticket #24659
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2353W" Then   'Let's Talk Science Ticket #27072 10/14/2015
        cmbCPP.AddItem "1"
        cmbCPP.AddItem "2"
    ElseIf glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
        cmbCPP.Clear
        cmbCPP.AddItem "0"
        cmbCPP.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
        cmbCPP.AddItem "1"
        cmbCPP.AddItem "2"
    ElseIf glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
        cmbCPP.AddItem "1"
        cmbCPP.AddItem "2"
    Else
        cmbCPP.AddItem "0"
        cmbCPP.AddItem "1"
        cmbCPP.AddItem "2"
        cmbCPP.AddItem "Y"
        cmbCPP.AddItem "N"
        cmbCPP.AddItem ""
    End If
End If

'Ticket #18188
If glbCompSerial = "S/N - 2418W" Then
    lblTitle(9).FontBold = True
End If

If glbCompSerial = "S/N - 2378W" Then 'Aurora
    
    'Ticket #20931 - Town of Aurora
    cmbUIC.AddItem "Y"
    cmbUIC.AddItem "N"
    
    'added by Bryan 25/Oct/05 Ticket#9607
    'cmbUIC.AddItem "0" 'Ticket #20931 - removed as per mapping document
    'cmbUIC.AddItem "1"
    'cmbUIC.AddItem ""   'Ticket #20931 - as per mapping document
    If glbVadim Then cmbUIC.ListIndex = 0
    
    
    txtGrossCalc.Text = "Y"     'Ticket #20931 - as per mapping document
    
'ElseIf glbInsync Then 'Frank Ticket# 7118
ElseIf glbCompSerial = "S/N - 2382W" Then 'Ticket #16235 Samuel
    cmbUIC.Width = 2500
    cmbUIC.AddItem "0 - Default EI Rate"
    cmbUIC.AddItem "2 - 2nd Mrs. Sam, Full EI"
    cmbUIC.AddItem "A - Combination P+S"
    cmbUIC.AddItem "B - Combination X+2"
    cmbUIC.AddItem "C - EI Weeks Recorded"
    cmbUIC.AddItem "E - Exempt from EI & WCB"
    cmbUIC.AddItem "N - Combination X+P"
    cmbUIC.AddItem "P - Full EI rate of 1.4"
    cmbUIC.AddItem "W - Combination Y+2"
    cmbUIC.AddItem "X - Exempt"
ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
    cmbUIC.Width = 2000
    cmbUIC.AddItem "0 - No Deductions"
    cmbUIC.AddItem "1 - Normal"
    cmbUIC.AddItem "5 - Commission"
    cmbUIC.AddItem "6 - Percent"
ElseIf glbCompSerial = "S/N - 2411W" Then   'Wellington-Dufferin-Guelph Public Health - Ticket #17129
    cmbUIC.AddItem "0"
    cmbUIC.AddItem "P"
ElseIf glbCompSerial = "S/N - 2436W" Then   'Family Day Care Services - Ticket #23603 Franks 04/30/2013
    cmbUIC.AddItem "0"
    cmbUIC.AddItem "P"
ElseIf glbInsync Then
    cmbUIC.AddItem "0"
    cmbUIC.AddItem "1"
    cmbUIC.AddItem "2"
    cmbUIC.AddItem "B"
    cmbUIC.AddItem "C"
    cmbUIC.AddItem "P"
    cmbUIC.AddItem "O"
    cmbUIC.AddItem "X"
    If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21519
        cmbUIC.Clear
        cmbUIC.AddItem "0" '"O" 'Ticket #22313
        cmbUIC.AddItem "X"
    End If
    'End If
   ' cmbUIC.AddItem ""
ElseIf glbVadim Then
    If glbCompSerial = "S/N - 2379W" Then   'Ticket #23795 - Town of LaSalle
        cmbUIC.AddItem "00"  'EI Code
        cmbUIC.AddItem "01"
        cmbUIC.AddItem "02"
    'Ticket #25412 - They do not need this - They have EI Rate.
    'ElseIf glbCompSerial = "S/N - 2447W" Then 'Or glbCompSerial = "S/N - 2458W" Then
    '    'Town of Greater Napanee - Ticket #24375
    '    'Ticket #24996 - City of Campbell River
    '    cmbUIC.AddItem "1"
    '    cmbUIC.AddItem ""
    Else
        cmbUIC.AddItem "Y"
        cmbUIC.AddItem "N"
        cmbUIC.ListIndex = 0
    End If
ElseIf glbCompSerial = "S/N - 2425W" Then 'Four Villages Ticket #21556
    cmbUIC.Width = 2500
    'Ticket #18221
    'cmbUIC.AddItem "1 - Deducts EI; ER rate 1.4"
    'cmbUIC.AddItem "2 - Exempt from EI; 1.4 rate"
    cmbUIC.AddItem "P - Deducts EI; ER rate 1.4"
    cmbUIC.AddItem "N - Exempt from EI; 1.4 rate"
    
ElseIf glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21519
    cmbUIC.AddItem "0" '"O" 'Ticket #21519 Franks 08/31/2012
    cmbUIC.AddItem "X"
Else
    If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
        cmbUIC.AddItem "01"  'EI Code
        cmbUIC.AddItem "02"
        cmbUIC.AddItem "03"
        cmbUIC.AddItem "11"
        cmbUIC.AddItem "12"
        cmbUIC.AddItem "13"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2241W" Then 'Granite Club
        cmbUIC.AddItem "FT"
        cmbUIC.AddItem "PT"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2385W" Then ' Conservation Halton 'Ticket #13063
        cmbUIC.AddItem "U1"
        cmbUIC.AddItem "U2"
        cmbUIC.AddItem "U3"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2381W" Then 'The Elliott Community  Ticket #13603
        cmbUIC.AddItem "Y"
        cmbUIC.AddItem "N"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2394W" Then 'St. John's  Ticket #15201
        cmbUIC.AddItem ""
        cmbUIC.AddItem "A"
        cmbUIC.AddItem "B"
        cmbUIC.AddItem "C"
        cmbUIC.AddItem "D"
        cmbUIC.AddItem "N"
        cmbUIC.AddItem "Y"
    ElseIf glbCompSerial = "S/N - 2379W" Then   'Town of LaSalle
        cmbUIC.AddItem "00"  'EI Code
        cmbUIC.AddItem "01"
        cmbUIC.AddItem "02"
    ElseIf glbCompSerial = "S/N - 2414W" Then   'Merrickville DCHSC 'Ticket #17210
        cmbUIC.AddItem "1"  'EI Code
        cmbUIC.AddItem "2"
    ElseIf glbCompSerial = "S/N - 2408W" Then   'Township of Wilmot - Ticket #19275
        cmbUIC.AddItem "Y"
        cmbUIC.AddItem "N"
'    ElseIf glbCompSerial = "S/N - 2425W" Then   'Four Villages Community Health Centre - Ticket #18221
'        cmbUIC.AddItem "0"
'        cmbUIC.AddItem "2"
'        cmbUIC.AddItem "B"
'        cmbUIC.AddItem "E"
'        cmbUIC.AddItem "J"
'        cmbUIC.AddItem "N"
'        cmbUIC.AddItem "P"
'        cmbUIC.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Autism Services  Ticket #20113
        cmbUIC.AddItem "1"
        cmbUIC.AddItem "2"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2396W" Then 'Oshawa CHC Ticket #20598 Franks 06/29/2012
        cmbUIC.AddItem "1"
        cmbUIC.AddItem "2"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2447W" Then   'Town of Greater Napanee - Ticket #22905
        cmbUIC.AddItem "Y"
        cmbUIC.AddItem "N"
        cmbUIC.AddItem ""
    ElseIf glbCompSerial = "S/N - 2457W" Then   'McLeod Law 'Ticket #24863
        'cmbUIC.AddItem "Y"
        'cmbUIC.AddItem "N"
        'Ticket #24864 Franks 06/10/2014
        cmbUIC.AddItem "0"
        cmbUIC.AddItem "2"
        cmbUIC.AddItem "B"
        cmbUIC.AddItem "E"
        cmbUIC.AddItem "J"
        cmbUIC.AddItem "N"
        cmbUIC.AddItem "P"
        cmbUIC.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
        cmbUIC.AddItem "0"
        cmbUIC.AddItem "2"
        cmbUIC.AddItem "B"
        cmbUIC.AddItem "E"
        cmbUIC.AddItem "J"
        cmbUIC.AddItem "N"
        cmbUIC.AddItem "P"
        cmbUIC.AddItem "X"
    ElseIf glbCompSerial = "S/N - 2453W" Then   'Town of Gander 'Ticket #24518 Franks 04/15/2014
        cmbUIC.Clear
        cmbUIC.AddItem "Y"
        cmbUIC.AddItem "N"
    ElseIf glbCompSerial = "S/N - 2353W" Then   'Let's Talk Science Ticket #27072 10/14/2015
        cmbUIC.AddItem "1"  'EI Code
        cmbUIC.AddItem "2"
    ElseIf glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
        cmbUIC.AddItem "1"  'EI Code
        cmbUIC.AddItem "2"
    ElseIf glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
        cmbUIC.AddItem "1"  'EI Code
        cmbUIC.AddItem "2"
    Else
        cmbUIC.AddItem "0"
        cmbUIC.AddItem "1"
        cmbUIC.AddItem "2"
        cmbUIC.AddItem "FT"
        cmbUIC.AddItem "PT"
        cmbUIC.AddItem ""
    End If
End If


'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    lblSupervisor.Caption = "Max. Bank Hours"
    clpCode(1).Tag = "00-Maximum Bank Hours"
End If
If glbWFC Then
    ''March 20, 2006 Bryan for Woodbridge
    'cboDepositCode.AddItem "W - Check 1"
    'cboDepositCode.AddItem "X - Check 2"
    'cboDepositCode.AddItem "Y - Saving 1"
    'cboDepositCode.AddItem "Z - Saving 2"
    'cboDepositCode2.AddItem "W - Check 1"
    'cboDepositCode2.AddItem "X - Check 2"
    'cboDepositCode2.AddItem "Y - Saving 1"
    'cboDepositCode2.AddItem "Z - Saving 2"
    'cboDepositCode3.AddItem "W - Check 1"
    'cboDepositCode3.AddItem "X - Check 2"
    'cboDepositCode3.AddItem "Y - Saving 1"
    'cboDepositCode3.AddItem "Z - Saving 2"
    'Oct 3, 2006 Frank for Woodbridge Ticket #11772
    cboDepositCode.AddItem "V - Check 1"
    cboDepositCode.AddItem "W - Check 2"
    cboDepositCode.AddItem "X - Saving 1"
    cboDepositCode.AddItem "Y - Saving 2"
    cboDepositCode2.AddItem "V - Check 1"
    cboDepositCode2.AddItem "W - Check 2"
    cboDepositCode2.AddItem "X - Saving 1"
    cboDepositCode2.AddItem "Y - Saving 2"
    cboDepositCode3.AddItem "V - Check 1"
    cboDepositCode3.AddItem "W - Check 2"
    cboDepositCode3.AddItem "X - Saving 1"
    cboDepositCode3.AddItem "Y - Saving 2"
    cboFedMarry.AddItem "S - Single"
    cboFedMarry.AddItem "M - Married"
    cboStateMarry.AddItem "S - Single"
    cboStateMarry.AddItem "M - Married"
    cboStateMarry.AddItem "? - Head of Household"
    'end bryan
End If

If glbCompSerial = "S/N - 2394W" Then 'St. John's  Ticket #15201
    cmbGrossCalc.AddItem "N"
    cmbGrossCalc.AddItem "Y"
    cmbGrossCalc.Left = txtGrossCalc.Left
    cmbGrossCalc.Visible = True
    txtGrossCalc.Visible = False
End If

If glbCompSerial = "S/N - 2411W" Then 'Wellington-Dufferin-Guelph Public Health - Ticket #17129
    cmbGrossCalc.AddItem "0"
    cmbGrossCalc.AddItem "M"
    cmbGrossCalc.Left = txtGrossCalc.Left
    cmbGrossCalc.Visible = True
    txtGrossCalc.Visible = False
End If

If glbCompSerial = "S/N - 2233W" Then   'Leeds-Grenville F&CS - Ticket #16737
    lblTitle(16).FontBold = True
    lblVadim1.FontBold = True
End If

If glbCompSerial = "S/N - 2382W" Then 'Ticket #18090 Samuel
    Call SamuelScreenSetup
End If

If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    lblUserText1.Caption = lStr("User Text 1")
    lblUserText1.Visible = True
    txtUserText1.Visible = True
End If

'Four Villages Community Health Centre - Ticket #18221
If glbCompSerial = "S/N - 2425W" Then
    lblTitle(9).FontBold = True
    lblTitle(10).FontBold = True 'E.I.
    lblTitle(19).FontBold = True 'C.P.P
End If
If glbCompSerial = "S/N - 2453W" Then  'Town of Gander Ticket #24518 Franks 06/04/2015
    lblTitle(9).FontBold = True 'Province of Employment
    lblTitle(10).FontBold = True 'E.I.
    lblTitle(19).FontBold = True 'C.P.P
    Label5.FontBold = True
    Label4.FontBold = True
End If

glbOnTop = "FRMEBANK"
MDIMain.panHelp(0).Caption = "Banking Information"

glbOnTop = "FRMEBANK"

If glbCountry <> "CANADA" Or glbEmpCountry = "U.S.A." Then
    lblTitle(1) = "Transit/ABA"
    lblTitle(2).Visible = False
    txtBankCode.Visible = False
    txtBankCode2.Visible = False
    txtBankCode3.Visible = False
    txtBranchCode.Visible = False
    txtBranchCode2.Visible = False
    txtBranchCode3.Visible = False
    txtTransitABA.Visible = True
    txtTransitABA2.Visible = True
    txtTransitABA3.Visible = True
    
    lblTitle(5) = "State Exemption"
    medTD1Amnt.Tag = "20-State Exemption"
    lblTitle(6) = "State Tax Code"
    txtTD1Code.Tag = "00-State Tax Code"
    lblTitle(7) = "State Extra Tax $"
    MedExtraTax.Tag = "20-State Extra Tax $"
    lblTitle(20) = "State Extra Tax %"
    medExtraTaxPC.Tag = "10-State Extra Tax %"
    lblTitle(11) = "Workers Comp Code"
    txtWCB.Tag = "30-Workers Comp Code"
    lblTitle(9) = "Worked State Tax Code"
     clpProv.Tag = "00-Worked State Tax Code"
    'Add by Franks on Jun 18,02 as Linda request - Begin
    Label4.Visible = False
    chkTD1Form.Visible = False
    Label5.Visible = False
    chkProvForm.Visible = False
    lblTitle(24).Visible = False
    lblTitle(25).Visible = False
    lblTitle(26).Visible = False
    lblTitle(27).Visible = False
    medProvAmt.Visible = False
    txtProvCode.Visible = False
    MedExtraTax.Visible = False
    medExtraTaxPC.Visible = False
    lblSupervisor.Visible = False
    clpCode(1).Visible = False
    lblCalcCode.Visible = False
    txtGrossCalc.Visible = False
    lblGarn.Visible = False
    txtGarn.Visible = False
    lblTitle(21).Visible = False
    medCHDSUP.Visible = False
    lblTitle(10).Visible = False
    cmbUIC.Visible = False
    lblTitle(11).Visible = False
    txtWCB.Visible = False
    lblTitle(12).Visible = False
    cmbPenCode.Visible = False
    lblTitle(19).Visible = False
    lblTitle(19).Caption = False
    cmbCPP.Visible = False
    lblTitle(35).Visible = False
    txtWSIBCde.Visible = False
    Label2.Visible = False
    txtPAYFREQ.Visible = False
    clpCode(1).ShowDescription = False
    'Add by Franks on Jun 18,02 as Linda request - End
    
    If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "U.S.A." Then   'I.T. Xchange
        lblTitle(12).Visible = True
        cmbPenCode.Visible = True
    End If
    
End If

If glbCompSerial = "S/N - 2417W" Then  'Ticket #22710 - County of Perth - Move Vacation Pay Percentage to where Salary Distribution was
    lblTitle(8).Visible = False
    medVacPPct.Visible = False
End If

If (glbCompSerial = "S/N - 2409W") Then lblSupervisor.FontBold = True 'Ticket #30066 Franks - Skylark Children

Call ComPenCode

Screen.MousePointer = DEFAULT
If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If
If glbLinamar Then Call ctrlSetup

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call addItems

If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    lblTitle(8).FontBold = True
End If
If glbCompSerial = "S/N - 2241W" Then 'Granite Club
    lblTitle(8).FontBold = True
    lblTitle(9).FontBold = True
    lblTitle(10).FontBold = True    'E.I Code
End If
If glbCompSerial = "S/N - 2454W" Then   'Showa Canada 'Ticket #24659
    lblTitle(8).FontBold = True
End If
If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA" Then   'I.T. Xchange
    lblTitle(9).FontBold = True
End If
If glbCompSerial = "S/N - 2375W" Then 'City of Timmins
    lblTitle(35).FontBold = True
    lblTitle(4).Caption = "Stat. Pay %"
    medPenPct.Tag = "10-Enter Stat. Pay %"
    'medPenPct.Format = ""
End If
If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
    lblGarn.Caption = "Max. Sick Accrual"
    txtGarn.Tag = "10-Enter Max. Sick Accrual"
    Label2.Caption = "Pay Vac. Every Period"
End If

If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership
    lblCalcCode.Caption = "Fed. Exemption Code"
    txtGrossCalc.Tag = "00-Fed. Exemption Code"
    cmbUIC.Visible = False
    txtUIC.Left = cmbUIC.Left
    txtUIC.Visible = True
    lblVadim1.FontBold = True
    
    If (NewHireForms.count > 0) Or fglbNew Then     'for New Hire only - Ticket #14634
        chkDirectDeposit = 1
    End If
End If

If glbCompSerial = "S/N - 2373W" Then 'District of Muskoka - Ticket #25670
    lblTitle(35).FontBold = True
    If (NewHireForms.count > 0) Or fglbNew Then     'New Hires
        cmbWSIBCode.ListIndex = 0
    End If
End If

'Ticket #27436 - Super Channel
If glbCompSerial = "S/N - 2475W" Then
    txtGrossCalc.Tag = "00-Tax Status - 1/X/F/M/A/I/P/N/V "
End If

frmGeneral.Top = 4080
frmGeneral.Left = 360

fraUSA.Top = 2160
fraUSA.Left = 360
fraUSA.Height = 3100 '2925 '2205
fraUSA.Width = 10215

fraOUTAddr.Top = 5010
fraOUTAddr.Left = 240
fraOUTAddr.Height = 1635
fraOUTAddr.Width = 9045

Call ST_UPD_MODE(False)
'If Not gSec_Upd_Banking Then
'    cmdModify.Enabled = False
'    cmdCancel.Enabled = False
'    cmdOK.Enabled = False
'End If
etGLfocus = False
Call INI_Controls(Me)
Call Display_Value

If glbCompSerial = "S/N - 2375W" Then 'City of Timmins
    If Len(clpProv) = 0 Then clpProv = "ON"
End If

If glbCompSerial = "S/N - 2408W" Then ''Township of Wilmot - Ticket #15785
    If Len(clpProv) = 0 Then clpProv = "ON"
End If

If glbCompSerial = "S/N - 2373W" Then 'District of Muskoka - Ticket #25670
    lblTitle(35).FontBold = True
    If (NewHireForms.count > 0) Or fglbNew Then     'New Hires
        cmbWSIBCode.ListIndex = 0
    End If
End If

If glbCompSerial = "S/N - 2458W" Then 'Ticket #25469 - City of Campbell River
    If (NewHireForms.count > 0) Or fglbNew Then     'New Hires
        cmbWSIBCode.ListIndex = 0
    End If
End If

'Ticket #28786 - Goodmans
If glbCompSerial = "S/N - 2290W" Then
    lblTitle(9).FontBold = True 'Province of Employment
    lblTitle(10).FontBold = True 'E.I.
    lblTitle(19).FontBold = True 'CPP
    lblTitle(11).Caption = "EI Pref."
End If

'County of Peterborough - Ticket #28993
'City of Kenora - Ticket #29054
If glbCompSerial = "S/N - 2486W" Or glbCompSerial = "S/N - 2487W" Then
    lblTitle(11).Caption = "WCB"
    lblTitle(35).Caption = "WCB Code"
    cmbWSIBCode.Visible = True
    txtWSIBCde.Visible = False
End If


If glbCompSerial = "S/N - 2487W" Then 'Ticket #30217 Franks 06/13/2017 City of Kenora
    Call ScreenSetupKenora
End If

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
If (glbEmpCountry = "CANADA") Then 'Ticket #15818
    If NewHireForms.count > 0 Then
    Dim Msg, DgDef, Response
        If Len(txtBankCode.Text) = 0 And Len(txtAccount.Text) = 0 Then
            Msg = "Banking information missed. It is required for Canadian employees." & Chr(10)
            Msg = Msg & "Are you sure you want to save anyway?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response = MsgBox(Msg, DgDef, "")
            If Response = IDNO Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End If
If glbWFC Then 'Ticket #17823 Mar 22, 2010
    If glbEmpCountry = "U.S.A." Then
            'Ticket #29828 Franks 02/14/2017 - don't use it since it is on Status/Date screen
            ''If Len(clpVadim21.Text) < 1 Then
            ''    MsgBox lStr("Vadim Field 2 is required field")
            ''    clpVadim21.SetFocus
            ''    Cancel = True
            ''    Exit Sub
            ''ElseIf Len(clpVadim21.Text) > 0 And clpVadim21.Caption = "Unassigned" Then
            ''    MsgBox lStr("Vadim Field 2 must be valid")
            ''    clpVadim21.SetFocus
            ''    Cancel = True
            ''    Exit Sub
            ''End If
        If Len(clpProvE.Text) < 1 Then
            MsgBox lblTitle(45) & " is required field"
            clpProvE.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
End If
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEBANK = Nothing  'carmen may 2000
Call NextForm
End Sub

Private Sub lblDirectDeposit_Change()

If lblDirectDeposit.Caption = "Y" Then
    chkDirectDeposit.Value = 1
Else
    chkDirectDeposit.Value = 0
End If

End Sub

Private Sub lblEEID_Change()

If Len(txtSurname.Text) > 0 And Len(txtFName.Text) > 0 Then  ' dont do on add new until in
    frmEBANK.Caption = "Banking Information - " & Left$(txtSurname, 5)
    frmEBANK.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If
If glbLinamar Then If Not IsNull(rsDATA("ED_OUTCOUNTRY")) Then comOUTCountry = rsDATA("ED_OUTCOUNTRY")
lblEENUM = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub lblProvForm_Change()
If lblProvForm.Caption = "Y" Then
   chkProvForm.Value = 1
Else
   chkProvForm.Value = 0
End If
End Sub

Private Sub lblQTBTORRSP_Change()
If lblQTBTORRSP.Caption = "Y" Then
   chkQTBTORRSP.Value = 1
Else
   chkQTBTORRSP.Value = 0
End If
End Sub

Private Sub lblTD1_Change()
If lblTD1.Caption = "Y" Then
   chkTD1Form.Value = 1
Else
   chkTD1Form.Value = 0
End If
End Sub

Private Sub medAmountDeposit_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medAmountDeposit2_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medAmountDeposit3_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medAmountDeposit_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medAmountDeposit_LostFocus()
    If medAmountDeposit = "" Then medAmountDeposit = 0
End Sub

Private Sub medAmountDeposit2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medAmountDeposit2_LostFocus()
    If medAmountDeposit2 = "" Then medAmountDeposit2 = 0
End Sub

Private Sub medAmountDeposit3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medAmountDeposit3_LostFocus()
    If medAmountDeposit3 = "" Then medAmountDeposit3 = 0
End Sub

Private Sub medCHDSUP_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub medCHDSUP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medCHDSUP_LostFocus()
If Len(Trim(medCHDSUP)) = 0 Then medCHDSUP = 0
End Sub

Private Sub medExtraTax_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub MedExtraTax_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub MedExtraTax_LostFocus()
If (Not IsNumeric(MedExtraTax)) Then MedExtraTax = 0  'And MedExtraTax.DataChanged Then MedExtraTax = 0

End Sub

Private Sub medExtraTaxPC_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medExtraTaxPC_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
'    If Len(medExtraTaxPC) > 0 Then
'        medExtraTaxPC = medExtraTaxPC * 100
'    End If
End Sub

Private Sub medExtraTaxPC_LostFocus()
If (Not IsNumeric(medExtraTaxPC)) And medExtraTaxPC.DataChanged Then medExtraTaxPC = 0
'If Len(medExtraTaxPC) > 0 Then
'    medExtraTaxPC = medExtraTaxPC / 100
'End If
End Sub

Private Sub medFedExtra_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medFedExtraPC_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medOUTPCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)

End Sub

Private Sub medOUTPCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub medPCDeposit_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medPCDeposit2_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medPCDeposit3_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medPCDeposit_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub
Private Sub medPCDeposit_LostFocus()
    If medPCDeposit = "" Then medPCDeposit = 0
End Sub

Private Sub medPCDeposit2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPCDeposit2_LostFocus()
    If medPCDeposit2 = "" Then medPCDeposit2 = 0
End Sub

Private Sub medPCDeposit3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPCDeposit3_LostFocus()
    If medPCDeposit3 = "" Then medPCDeposit3 = 0
End Sub

Private Sub medPenPct_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
    If Len(medPenPct) > 0 Then
        medPenPct = medPenPct * 100
    End If
End Sub

Private Sub medPenPct_KeyPress(KeyAscii As Integer)
If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medPenPct_LostFocus()
    If Len(medPenPct) > 0 Then
        medPenPct = medPenPct / 100
    End If
End Sub

Private Sub medProvAmt_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medProvAmt_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medProvAmt_LostFocus()
If (Not IsNumeric(medProvAmt.Text) Or medProvAmt.Text = "") Then medProvAmt = 0        'And medProvAmt.DataChanged Then medProvAmt = 0
End Sub

Private Sub medStateExtra_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medStateExtraPC_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTD1Amnt_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medTD1Amnt_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTD1Amnt_LostFocus()
    If (Not IsNumeric(medTD1Amnt)) Then medTD1Amnt = 0 'And medTD1Amnt.DataChanged Then medTD1Amnt = 0
End Sub

Private Sub medTD3_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medTD3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTD3_LostFocus()
    If (Not IsNumeric(medTD3)) Then medTD3 = 0     'And medTD3.DataChanged Then medTD3 = 0
End Sub

Private Sub medTD3PC_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medTD3PC_GotFocus()
Call SetPanHelp(Me.ActiveControl)
'Comment by Frank Dec 2,2003, ED_TD3PC is integer, it can't be saved as a number less 1
'If Len(medTD3PC) > 0 Then
'    medTD3PC = medTD3PC * 100
'End If
End Sub

Private Sub medTD3PC_LostFocus()
If (Not IsNumeric(medTD3PC)) And medTD3PC.DataChanged Then medTD3PC = 0
'Comment by Frank Dec 2,2003, ED_TD3PC is integer, it can't be saved as a number less 1
'If Len(medTD3PC) > 0 Then
'    medTD3PC = medTD3PC / 100
'End If
End Sub

Private Sub medVacPPct_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medVacPPct_GotFocus()
Call SetPanHelp(Me.ActiveControl)
If Len(medVacPPct) > 0 Then
    medVacPPct = medVacPPct * 100
End If
End Sub

Private Sub medVacPPct_LostFocus()
If (Not IsNumeric(medVacPPct)) And medVacPPct.DataChanged Then medVacPPct = 0
If Len(medVacPPct) > 0 Then
    medVacPPct = medVacPPct / 100
End If
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

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdPrint.Enabled = FT
'cmdModify.Enabled = FT
'cmdClose.Enabled = FT

'chkDirectDeposit.Enabled = TF
chkTD1Form.Enabled = TF
cmbCPP.Enabled = TF
cmbPenCode.Enabled = TF
cmbUIC.Enabled = TF
fraOUTAddr.Enabled = TF
'If glbWFC Then 'Ticket #22776 Franks 11/08/2012, WFC needs this enable
'    'On Bank 1, do not allow them to enter "Amount Deposit" and "Percent Deposited".
'    '02/25/10
'    medAmountDeposit.Enabled = False
'    medPCDeposit.Enabled = False
'Else
    medAmountDeposit.Enabled = TF
    medPCDeposit.Enabled = TF
'End If
medAmountDeposit2.Enabled = TF
medAmountDeposit3.Enabled = TF
medPCDeposit2.Enabled = TF
medPCDeposit3.Enabled = TF
medTD1Amnt.Enabled = TF
medTD3.Enabled = TF
medVacPPct.Enabled = TF
txtAccount.Enabled = TF
txtAccount2.Enabled = TF
txtAccount3.Enabled = TF
txtBankCode.Enabled = TF
txtBankCode2.Enabled = TF
txtBankCode3.Enabled = TF
txtBranchCode.Enabled = TF
txtBranchCode2.Enabled = TF
txtBranchCode3.Enabled = TF
clpCode(1).Enabled = TF
txtDepositCode.Enabled = TF
txtDepositCode2.Enabled = TF
txtDepositCode3.Enabled = TF
txtGarn.Enabled = TF
txtGrossCalc.Enabled = TF
 clpProv.Enabled = TF
txtTD1Code.Enabled = TF
txtWCB.Enabled = TF
medTD3PC.Enabled = TF
medCHDSUP.Enabled = TF
txtWSIBCde.Enabled = TF
txtPAYFREQ.Enabled = TF
txtFedTax.Enabled = TF
txtExtAmt.Enabled = TF
chkProvForm.Enabled = TF
medProvAmt.Enabled = TF
MedExtraTax.Enabled = TF
medExtraTaxPC.Enabled = TF
txtProvCode.Enabled = TF
chkDirectDeposit.Enabled = TF
clpVadim1.Enabled = TF
clpVadim2.Enabled = TF
cmbWCB.Enabled = TF
cmbWSIBCode.Enabled = TF
If glbLinamar Then
    txtExtrAnn.Enabled = TF
    chkQTBTORRSP.Enabled = TF
    cmbToPayroll.Enabled = TF
End If
If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership
    cmbUIC.Visible = False
    lblSupervisor.Enabled = False
    clpCode(1).Enabled = False
    lblGarn.Enabled = False
    txtGarn.Enabled = False
    lblTitle(21).Enabled = False
    medCHDSUP.Enabled = False
    lblVadim2.Enabled = False
    clpVadim2.Enabled = False
    lblTitle(11).Enabled = False
    cmbWCB.Enabled = False
    lblTitle(12).Enabled = False
    cmbPenCode.Enabled = False
    lblTitle(35).Enabled = False
    cmbWSIBCode.Enabled = False
    Label2.Enabled = False
    cmbPayFreq.Enabled = False
    lblTitle(4).Enabled = False
    medPenPct.Enabled = False
    txtWCB.Enabled = False
    txtWSIBCde.Enabled = False
    txtPAYFREQ.Enabled = False
End If
If glbCompSerial = "S/N - 2385W" Then 'Conservation Halton - Ticket #13165
    lblSupervisor.Enabled = False
    clpCode(1).Enabled = False
    lblGarn.Enabled = False
    txtGarn.Enabled = False
    lblTitle(21).Enabled = False
    medCHDSUP.Enabled = False
    lblVadim1.Enabled = False
    clpVadim1.Enabled = False
    lblVadim2.Enabled = False
    clpVadim2.Enabled = False
    cmbWCB.Enabled = False
    Label2.Enabled = False
    cmbPayFreq.Enabled = False
    lblTitle(4).Enabled = False
    medPenPct.Enabled = False
    lblCalcCode.Enabled = False
    txtGrossCalc.Enabled = False
    txtPAYFREQ.Enabled = False
End If
If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    txtUserText1.Enabled = TF
End If
End Sub

Private Sub txtAccount_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAccount2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAccount3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBankCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBankCode_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If glbCountry = "CANADA" Then If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBankCode_LostFocus()
    'New Bank Data enter then default Direct Deposit On
    If Len(txtBankCode) > 0 And Len(txtAccount) = 0 Then
        chkDirectDeposit.Value = 1
        Call chkDirectDeposit_Click
    End If
End Sub

Private Sub txtBankCode2_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBankCode3_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBankCode2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBankCode3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBranchCode_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBranchCode2_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBranchCode3_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtBranchCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBranchCode2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBranchCode3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub txtCPP_Change()
    If glbInsync Then
        If glbCompSerial = "S/N - 2382W" Then
            Call To_cmbCPP_index(txtCPP.Text)
        ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
            Call To_cmbCPP_index(txtCPP.Text)
        Else
            If txtCPP = "" Or txtCPP = "0" Then
                If glbCompSerial = "S/N - 2292W" And txtCPP = "" Then   'County of Elgin
                    cmbCPP.ListIndex = -1
                Else
                    cmbCPP.ListIndex = 0
                End If
            Else
                If glbCompSerial = "S/N - 2292W" And cmbCPP.ListIndex = -1 Then   'County of Elgin
                    If txtCPP = "X" Then
                        cmbCPP.ListIndex = 1
                    Else
                        cmbCPP.ListIndex = -1
                    End If
                Else
                    cmbCPP.ListIndex = 1
                End If
            End If
        End If
    Else
        cmbCPP = txtCPP
    End If
End Sub

Private Sub To_cmbCPP_index(xCode)
    If glbCompSerial = "S/N - 2382W" Then
        Select Case xCode
            Case "0": cmbCPP.ListIndex = 0
            Case "E": cmbCPP.ListIndex = 1
            Case "M": cmbCPP.ListIndex = 2
            Case "X": cmbCPP.ListIndex = 3
        End Select
    End If
    If glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
        Select Case Trim(xCode)
            Case "": cmbCPP.ListIndex = 0
            Case "X": cmbCPP.ListIndex = 1
            'Ticket #22667
            'Case Else
            '    cmbCPP.ListIndex = -1
        End Select
    End If
End Sub

Private Sub txtCPP_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDepositCode_Change()
    cboDepositCode.Text = txtDepositCode.Text
End Sub
Private Sub txtDepositCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDepositCode_LostFocus()
    txtDepositCode.Text = UCase(txtDepositCode.Text)
End Sub

Private Sub txtDepositCode2_Change()
    cboDepositCode2.Text = txtDepositCode2.Text
End Sub
Private Sub txtDepositCode2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDepositCode2_LostFocus()
    txtDepositCode2.Text = UCase(txtDepositCode2.Text)
End Sub

Private Sub txtDepositCode3_Change()
    cboDepositCode3.Text = txtDepositCode3.Text
End Sub
Private Sub txtDepositCode3_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDepositCode3_LostFocus()
    txtDepositCode3.Text = UCase(txtDepositCode3.Text)
End Sub

Private Sub txtExtAmt_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtExtAmt_LostFocus()
If (Not IsNumeric(txtExtAmt)) And txtExtAmt.DataChanged Then txtExtAmt = 0

End Sub

Private Sub txtExtrAnn_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtExtrAnn_LostFocus()
    If (Not IsNumeric(txtExtrAnn)) And txtExtrAnn.DataChanged Then txtExtrAnn = 0
End Sub

Private Sub medFedTax_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub



Private Sub txtFedExemp_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtFedMarry_Change()
     cboFedMarry.Text = txtFedMarry.Text
End Sub

Private Sub txtFedTax_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtFedTax_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtFedTax_LostFocus()
If (Not IsNumeric(txtFedTax)) And txtFedTax.DataChanged Then txtFedTax = 0

End Sub

Private Sub txtGarn_KeyPress(KeyAscii As Integer)
    ' AC - dkostka - 05/08/2001 - Only let the user enter numbers
    If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtGarn_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtGarn_LostFocus()
    If Len(Trim(txtGarn)) = 0 Then txtGarn = 0
End Sub

Private Sub txtGrossCalc_Change()
'Wellington-Dufferin-Guelph Public Health - Ticket #17129
If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2411W" Then 'St. John's  Ticket #15201
    cmbGrossCalc = txtGrossCalc
End If
End Sub

Private Sub txtGrossCalc_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtGrossCalc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtOUTAddr_GotFocus()
    Call SetPanHelp(Me.ActiveControl)

End Sub


Private Sub txtPAYFREQ_Change()
    If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
        If txtPAYFREQ = "Y" Then
            cmbPayFreq.ListIndex = 0
        ElseIf txtPAYFREQ = "N" Then
            cmbPayFreq.ListIndex = 1
        Else
            cmbPayFreq.ListIndex = -1
        End If
    End If
End Sub

Private Sub txtPAYFREQ_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPension_Change()
    If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Then   'City of Kawartha Lakes & City of Timmins
        If txtPension.Text = "1" Then
            cmbPenCode.ListIndex = 0
        ElseIf txtPension.Text = "2" Then
            cmbPenCode.ListIndex = 1
        End If
    Else
        cmbPenCode.Text = txtPension.Text
    End If
End Sub

Private Sub txtPension_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub
'Private Sub txtProv_Change()
'If Prov_Snap.State = 0 Then Exit Sub
'Call ProvN_Desc
'End Sub
'Private Sub txtProv_DblClick()
'Call Get_Prov(False)
'If Len(glbProv) > 0 Then
'    txtProv.Text = glbProv
'    lblProvDesc.Caption = glbProvDesc
'End If
'Call ProvN_Desc
'End Sub
'Private Sub txtProv_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub
'Private Sub txtProv_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub
'Private Sub txtProv_LostFocus()
'    txtProv = UCase(txtProv)
'    Call ProvN_Desc
'End Sub
'Private Sub txtProv_Validate(Cancel As Boolean)
'Call ProvN_Desc
'End Sub
Private Sub txtProvCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtStateExemption_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtStateMarry_Change()
    cboStateMarry.Text = txtStateMarry.Text
End Sub

Private Sub txtStatusFalg3_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtTD1Code_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub txtToPayroll_Change()
cmbToPayroll.Text = txtToPayroll.Text
End Sub

Private Sub txtUIC_Change()
    If glbCompSerial = "S/N - 2382W" Then
        Call To_cmbUIC_index(txtUIC.Text)
    ElseIf glbCompSerial = "S/N - 2383W" Then  'Town of Orangeville Ticket #21207 Franks 11/15/2011
        Call To_cmbUIC_index(txtUIC.Text)
    ElseIf glbCompSerial = "S/N - 2425W" Then  'Four Villages Ticket #21556
        Call To_cmbUIC_index(txtUIC.Text)
    Else
        cmbUIC.Text = txtUIC.Text
    End If
End Sub

Private Sub To_cmbUIC_index(xUICCode)
    If glbSamuel Then
        Select Case Trim(xUICCode)
            Case "0": cmbUIC.ListIndex = 0
            Case "2": cmbUIC.ListIndex = 1
            Case "A": cmbUIC.ListIndex = 2
            Case "B": cmbUIC.ListIndex = 3
            Case "C": cmbUIC.ListIndex = 4
            Case "E": cmbUIC.ListIndex = 5
            Case "N": cmbUIC.ListIndex = 6
            Case "P": cmbUIC.ListIndex = 7
            Case "W": cmbUIC.ListIndex = 8
            Case "X": cmbUIC.ListIndex = 9
        End Select
     End If
     If glbCompSerial = "S/N - 2383W" Then  'Town of Orangeville Ticket #21207 Franks 11/15/2011
        Select Case Trim(xUICCode)
            Case "0": cmbUIC.ListIndex = 0
            Case "1": cmbUIC.ListIndex = 1
            Case "5": cmbUIC.ListIndex = 2
            Case "6": cmbUIC.ListIndex = 3
            'Ticket #22667
            Case Else
                cmbUIC.ListIndex = -1
        End Select
     End If
     If glbCompSerial = "S/N - 2425W" Then  'Four Villages Ticket #21556
        Select Case Trim(xUICCode)
            'Ticket #18221
            'Case "1": cmbUIC.ListIndex = 0
            'Case "2": cmbUIC.ListIndex = 1
            Case "P": cmbUIC.ListIndex = 0
            Case "N": cmbUIC.ListIndex = 1
            Case Else
                cmbUIC.ListIndex = -1
        End Select
     End If
End Sub

Private Sub txtUIC_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtWCB_Change()
 If glbInsync Then
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18417
        Call setcmbWCBListInd_Samuel
    ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21207 Franks 11/15/2011
        Call setcmbWCBListInd_Orangeville
    Else
        'Ticket #18437 May 7, 2010 Frank
        'If txtWCB = "" Or txtWCB = "0" Then
        '    cmbWCB.ListIndex = 0
        'Else
        '    cmbWCB.ListIndex = 1
        'End If
        Call setcmbWCBListInd_Comm
    End If
 End If
 If glbVadim Then
    'Town of Lasalle - not for them
    If glbCompSerial <> "S/N - 2379W" Then
        cmbWCB.ListIndex = Val(txtWCB) - 1
    End If
 End If
If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    cmbWCB.Text = txtWCB.Text
End If
If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics - Ticket #22716
    cmbWCB.Text = txtWCB.Text
End If
If glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
    cmbWCB.Text = txtWCB.Text
End If
If glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
    cmbWCB.Text = txtWCB.Text
End If
End Sub
Private Sub setcmbWCBListInd_Comm()
    Select Case txtWCB.Text
    Case "0": cmbWCB.ListIndex = 0
    Case "M": cmbWCB.ListIndex = 1
    Case "O": cmbWCB.ListIndex = 2
    Case "P": cmbWCB.ListIndex = 3
    Case "X": cmbWCB.ListIndex = 4
    Case Else: cmbWCB.ListIndex = -1
    End Select
End Sub
Private Sub setcmbWCBListInd_Samuel()
    Select Case txtWCB.Text
    Case "0": cmbWCB.ListIndex = 0
    Case "A": cmbWCB.ListIndex = 1
    Case "B": cmbWCB.ListIndex = 2
    Case "C": cmbWCB.ListIndex = 3
    Case "F": cmbWCB.ListIndex = 4
    Case "H": cmbWCB.ListIndex = 5
    Case "I": cmbWCB.ListIndex = 6
    Case "M": cmbWCB.ListIndex = 7
    Case "N": cmbWCB.ListIndex = 8
    Case "O": cmbWCB.ListIndex = 9
    Case "P": cmbWCB.ListIndex = 10
    Case "Q": cmbWCB.ListIndex = 11
    Case "T": cmbWCB.ListIndex = 12
    Case "X": cmbWCB.ListIndex = 13
    Case Else: cmbWCB.ListIndex = -1
    End Select
End Sub

Private Sub setcmbWCBListInd_Orangeville()
    Select Case Trim(txtWCB.Text)
    Case "": cmbWCB.ListIndex = 0
    Case "X": cmbWCB.ListIndex = 1
    Case Else: cmbWCB.ListIndex = -1
    End Select
End Sub

Private Sub txtWCB_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
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


Private Sub txtWCB_KeyPress(KeyAscii As Integer)
If glbPayWeb Then
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End If
End Sub

Private Sub txtWSIBCde_Change()
    Dim c As Long
    If glbCompSerial = "S/N - 2382W" Then 'Ticket #18417 Samuel
        Select Case txtWSIBCde.Text
        Case ""
            cmbWSIBCode.ListIndex = -1
        Case "0"
            cmbWSIBCode.ListIndex = 0
        Case "X"
            cmbWSIBCode.ListIndex = 1
        End Select
    Else
        If glbVadim And txtWSIBCde <> "0" Then
            For c = 0 To cmbWSIBCode.ListCount - 1
                If cmbWSIBCode.ItemData(c) = Val(txtWSIBCde.Text) Then
                    cmbWSIBCode.ListIndex = c
                    doOnce = True
                    Exit For
                End If
            Next c
            If txtWSIBCde = "" Then
                cmbWSIBCode.ListIndex = -1
            End If
        ElseIf glbCompSerial = "S/N - 2486W" Then   'County of Peterborough - Ticket #28993
            Call PopWSIBCodeCountyeterborough
        ElseIf glbCompSerial = "S/N - 2487W" Then   'City of Kenora - Ticket #29054
            cmbWSIBCode.Text = txtWSIBCde.Text
        End If
    End If
End Sub

Private Sub PopWSIBCodeCountyeterborough()
    'cmbWSIBCode.Text = txtWSIBCde.Text
    Select Case txtWSIBCde.Text
    Case ""
        cmbWSIBCode.ListIndex = -1
    Case "590-1"
        cmbWSIBCode.ListIndex = 0
    Case "590-2"
        cmbWSIBCode.ListIndex = 1
    Case "817-1"
        cmbWSIBCode.ListIndex = 2
    Case "817-2"
        cmbWSIBCode.ListIndex = 3
    Case "845-1"
        cmbWSIBCode.ListIndex = 4
    Case "845-2"
        cmbWSIBCode.ListIndex = 5
    Case "COM-1"
        cmbWSIBCode.ListIndex = 6
    Case "COM-2"
        cmbWSIBCode.ListIndex = 7
    End Select
        
End Sub

Private Sub txtWSIBCde_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_DDI,"
SQLQ = SQLQ & "ED_DEPOSIT, ED_BANK, ED_BRANCH, ED_ACCOUNT,"
SQLQ = SQLQ & "ED_AMTDEPOSIT, ED_PCDEPOSIT, ED_DEPOSIT2,"
SQLQ = SQLQ & "ED_BANK2, ED_BRANCH2, ED_ACCOUNT2,"
SQLQ = SQLQ & "ED_AMTDEPOSIT2, ED_PCDEPOSIT2, ED_DEPOSIT3,"
SQLQ = SQLQ & "ED_BANK3, ED_BRANCH3, ED_ACCOUNT3,"
SQLQ = SQLQ & "ED_AMTDEPOSIT3, ED_PCDEPOSIT3, ED_TD1, ED_TD1DOL,"
SQLQ = SQLQ & "ED_TD1CODE, ED_FEDTAX, ED_TD3, ED_TD3PC,"
SQLQ = SQLQ & "ED_EXTAMT, ED_PROVFORM, ED_PROVAMT, ED_PROVCODE,"
SQLQ = SQLQ & "ED_EXTRATAX, ED_EXTRATAXPC, ED_VACPC, ED_PROVEMP,"
SQLQ = SQLQ & "ED_SUPCODE, ED_GROSSCD, ED_GARN, ED_CHDSUP,"
SQLQ = SQLQ & "ED_WCB, ED_UIC, ED_PENSION, ED_CPP, ED_WCBCODE,"
SQLQ = SQLQ & "ED_TRANSITABA,ED_TRANSITABA2,ED_TRANSITABA3,"
SQLQ = SQLQ & "ED_PAYROLL_ID,ED_VADIM1,ED_VADIM2,"
SQLQ = SQLQ & "ED_PAYFREQ,ED_OUTADDR,ED_OUTPROV,ED_OUTCITY,ED_OUTCOUNTRY,ED_OUTPCODE,ED_OUTADDRT4,"
SQLQ = SQLQ & "ED_PENPCT,ED_DOH,"
SQLQ = SQLQ & "ED_USER_TEXT1," 'Ticket #28515 Franks 04/26/2016
If glbLinamar Then
    SQLQ = SQLQ & "ED_EXTRANN,ED_QTBTORRSP,"
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/26/2011
    SQLQ = SQLQ & "ED_PENPCTFIXED,"
End If
SQLQ = SQLQ & "ED_LDATE, ED_LTIME, ED_LUSER, ED_COUNTRY"
'City of Timmins - For RPP # (Vadim)
If glbCompSerial = "S/N - 2375W" Then
    SQLQ = SQLQ & ",ED_NORMALR, ED_DOB"
End If
If glbWFC Then 'Ticket #22553 Franks 09/18/2012
    'SQLQ = SQLQ & ",ED_COMBINATION"
    'Ticket #22584 Franks 09/27/2012 - use ED_HOMEWRKCNT instead of ED_COMBINATION
    SQLQ = SQLQ & ",ED_HOMEWRKCNT"
End If
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ
End Function

Private Sub ctrlSetup()
lblTitle(13).Visible = False
txtDepositCode.Visible = False
txtDepositCode2.Visible = False
txtDepositCode3.Visible = False
lblTitle(15).Visible = False
medPCDeposit.Visible = False
medPCDeposit2.Visible = False
medPCDeposit3.Visible = False
lblTitle(10).Visible = False
cmbUIC.Visible = False
frmlinamar.Visible = True
fraOUTAddr.Visible = True

frmGeneral.Visible = False
End Sub

' AC - dkostka - 05/08/2001 - Added function to find out if an ascii character is valid
' for a numeric entry field or not.
Private Function IsNumericEntry(KeyAscii As Integer, Optional NegAllowed As Boolean) As Boolean
    If KeyAscii = Asc(vbBack) Or IsNumeric(Chr(KeyAscii)) Or (NegAllowed And KeyAscii = Asc("-")) Then IsNumericEntry = True
End Function



Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comOUTCountry) = 0 And comOUTCountry <> "" Then
    xCountryList = xCountryList & "&" & comOUTCountry
    comOUTCountry.AddItem comOUTCountry
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub addItems()
Dim ctylist, x
ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        comOUTCountry.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        comOUTCountry.AddItem ctylist
    End If
Loop
End Sub
''' Sam add July 2002 * Remove ADO
Public Sub Display_Value()
    Dim SQLQ
    doOnce = False
    If rsDATA.EOF Or rsDATA.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If
    
  
    If glbtermopen Then
        SQLQ = "Select " & FldList & " from Term_HREMP"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockPessimistic
        
    Else
        SQLQ = "Select " & FldList & " from HREMP "
        SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    End If
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    glbEmpCountry = rsDATA("ED_COUNTRY")
    Call CANADA_USA_Banking
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    'County of Peterborough - Ticket #28993
    'City of Kenora - Ticket #29054
    If glbCompSerial = "S/N - 2486W" Or glbCompSerial = "S/N - 2487W" Then
        lblTitle(11).Caption = "WCB"
        lblTitle(35).Caption = "WCB Code"
    End If
    
    Me.cmdModify_Click
 End Sub


Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
    If chkOUTADDRT4.Value = Null Then chkOUTADDRT4.Value = False
End If
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Banking
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
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/20/2014
    TF = getFamilyDayUpdateRight(UpdateRight, glbLEE_ID)
Else
    If Not UpdateRight Then TF = False
End If
Call ST_UPD_MODE(TF)
End Sub

Private Function modEmpNoUpdate()
Dim dyn_Table As New ADODB.Recordset
Dim xCount, xx
Dim SQLQ, x%, xFldTitle, xFld As String, xTable As String
modEmpNoUpdate = False
On Error GoTo modUpdate_cmdUpdErr
Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
SQLQ = "SELECT * FROM INFO_HR_TABLES WHERE TERMINATION_TABLE=0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

dyn_Table.Open SQLQ, gdbAdoIhr001, adOpenStatic
MDIMain.panHelp(0).FloodPercent = 10
xCount = dyn_Table.RecordCount
xx = 0
Do Until dyn_Table.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 60 + 10
    xTable = dyn_Table("Table_Name")
    If IsNull(dyn_Table("EMPNBR_Alias")) Then xFld = "" Else xFld = dyn_Table("EMPNBR_Alias")
    If InStr(xFld, "_") = 0 Then xFldTitle = "" Else xFldTitle = Left(xFld, 3)
    If dyn_Table("Employee_Keyed") Then
        If xTable = "HREEO" Then
            Call UpdateLUSER(xTable, xFldTitle & "EMPNBR")
        Else
            Call UpdateEMPNBR(xTable, xFld, xFldTitle)
            Call UpdateLUSER(xTable, xFldTitle & "LUSER")
        End If
        Select Case xTable
        Case "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPER", xFldTitle)
        Case "HR_JOB_HISTORY", "HR_PERFORM_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU2", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU3", xFldTitle)
        Case "HR_OCC_HEALTH_SAFETY"
            Call UpdateEMPNBR(xTable, xFldTitle & "EMPNOT", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPERVISOR", xFldTitle)
        End Select
    Else
        If Len(xFld) > 0 Then Call UpdateLUSER(xTable, xFld)
    End If
    dyn_Table.MoveNext
    xx = xx + 1
Loop
MDIMain.panHelp(0).FloodPercent = 70


'Franks May 09,2002 for Essex County Library
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodPercent = 0
Screen.MousePointer = DEFAULT
modEmpNoUpdate = True

Exit Function
modUpdate_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MDIMain.panHelp(0).FloodType = 0
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdate Error", xTable, "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub SPCUpdatePayPeriod(xEmpnbr, xPayPeriod)
    Dim SQLQ As String
    
    SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_PAYP = '" & xPayPeriod & "'"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpnbr & " AND SH_CURRENT <>0"
    gdbAdoIhr001.Execute SQLQ

End Sub

Private Sub UpdateLUSER(nTable As String, nFld As String)
' dkostka - 10/11/2001 - Don't update LUSER, this is set to USERID now, which doesn't change
'   with employee number.

'Dim SQLQ
'SQLQ = "UPDATE " & ntable & " SET "
'SQLQ = SQLQ & nFld & "=" & getEmpnbr(txtEmpNum(1))
'SQLQ = SQLQ & " WHERE " & nFld & "= " & getEmpnbr(txtEmpNum(0))
'gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub UpdateEMPNBR(nTable As String, nFld As String, nFldTitle)
Dim SQLQ
SQLQ = "UPDATE " & nTable & " SET "
SQLQ = SQLQ & nFld & "=" & getEmpnbr(glbSPCNewEmpNo) & ","

If nFldTitle = "CR_" Or nFldTitle = "CT_" Or nFldTitle = "RC_" Then
    SQLQ = SQLQ & nFldTitle & "LTime = '" & Time$ & "',"
    SQLQ = SQLQ & nFldTitle & "LDate = " & Date_SQL(Date) & ","
Else
    SQLQ = SQLQ & nFldTitle & "LTIME = '" & Time$ & "',"
    SQLQ = SQLQ & nFldTitle & "LDATE = " & Date_SQL(Date) & ","
End If
SQLQ = SQLQ & nFldTitle & "LUSER = '" & glbUserID & "'"
SQLQ = SQLQ & " WHERE "
SQLQ = SQLQ & nFld & "= " & getEmpnbr(lblEEID) '
gdbAdoIhr001.Execute SQLQ
End Sub


Private Sub CANADA_USA_Banking()
    
    'If glbCountry <> "CANADA" Or glbEmpCountry = "U.S.A." Then
    If glbCountry = "U.S.A." Or glbEmpCountry = "U.S.A." Then
        lblTitle(1) = "Transit/ABA"
        lblTitle(2).Visible = False
        txtBankCode.Visible = False
        txtBankCode2.Visible = False
        txtBankCode3.Visible = False
        txtBranchCode.Visible = False
        txtBranchCode2.Visible = False
        txtBranchCode3.Visible = False
        txtTransitABA.Visible = True
        txtTransitABA2.Visible = True
        txtTransitABA3.Visible = True
        
        lblTitle(5) = "State Exemption"
        medTD1Amnt.Tag = "20-State Exemption"
        lblTitle(6) = "State Tax Code"
        txtTD1Code.Tag = "00-State Tax Code"
        lblTitle(7) = "State Extra Tax $"
        MedExtraTax.Tag = "20-State Extra Tax $"
        lblTitle(20) = "State Extra Tax %"
        medExtraTaxPC.Tag = "10-State Extra Tax %"
        lblTitle(11) = "Workers Comp Code"
        txtWCB.Tag = "30-Workers Comp Code"
        lblTitle(9) = "Worked State Tax Code"
         clpProv.Tag = "00-Worked State Tax Code"
        Label4.Visible = False
        chkTD1Form.Visible = False
        Label5.Visible = False
        chkProvForm.Visible = False
        lblTitle(24).Visible = False
        lblTitle(25).Visible = False
        lblTitle(26).Visible = False
        lblTitle(27).Visible = False
        medProvAmt.Visible = False
        txtProvCode.Visible = False
        MedExtraTax.Visible = False
        medExtraTaxPC.Visible = False
        lblSupervisor.Visible = False
        clpCode(1).Visible = False
        lblCalcCode.Visible = False
        txtGrossCalc.Visible = False
        lblGarn.Visible = False
        txtGarn.Visible = False
        lblTitle(21).Visible = False
        medCHDSUP.Visible = False
        lblTitle(10).Visible = False
        cmbUIC.Visible = False
        lblTitle(11).Visible = False
        txtWCB.Visible = False
        lblTitle(12).Visible = False
        cmbPenCode.Visible = False
        lblTitle(35).Visible = False
        txtWSIBCde.Visible = False
        Label2.Visible = False
        txtPAYFREQ.Visible = False
        clpCode(1).ShowDescription = False
        cmbCPP.Visible = False
        lblTitle(19).Visible = False
        lblTitle(10).Visible = False
        cmbUIC.Visible = False
        lblVadim1.Visible = False
        lblVadim2.Visible = False
        clpVadim1.Visible = False
        clpVadim2.Visible = False
    Else
        lblTitle(1) = "Bank Code"
        lblTitle(2).Visible = True
        txtBankCode.Visible = True
        txtBankCode2.Visible = True
        txtBankCode3.Visible = True
        txtBranchCode.Visible = True
        txtBranchCode2.Visible = True
        txtBranchCode3.Visible = True
        txtTransitABA.Visible = False
        txtTransitABA2.Visible = False
        txtTransitABA3.Visible = False
        
        lblTitle(5) = "TD1 Amount"
        medTD1Amnt.Tag = "20-Amount as found on TD1"
        lblTitle(6) = "TD1 Code"
        txtTD1Code.Tag = "00-TD1 code as reported on the TD1 Form"
        lblTitle(7) = "Extra Tax"
        MedExtraTax.Tag = "20-Extra Tax on Provincial Form"
        lblTitle(20) = "Extra Tax %"
        medExtraTaxPC.Tag = "10-Extra Tax Percentage on Provincial Form"
        lblTitle(11) = "W.S.I.B."
        txtWCB.Tag = "00-Workers Compensation Board "
        lblTitle(9) = "Province of Employment"
        clpProv.Tag = "30-Province Code"
        Label4.Visible = True
        chkTD1Form.Visible = True
        Label5.Visible = True
        chkProvForm.Visible = True
        lblTitle(24).Visible = True
        lblTitle(25).Visible = True
        lblTitle(26).Visible = True
        lblTitle(27).Visible = True
        medProvAmt.Visible = True
        txtProvCode.Visible = True
        MedExtraTax.Visible = True
        medExtraTaxPC.Visible = True
        lblSupervisor.Visible = True
        clpCode(1).Visible = True
        lblCalcCode.Visible = True
        txtGrossCalc.Visible = True
        lblGarn.Visible = True
        txtGarn.Visible = True
        lblTitle(21).Visible = True
        medCHDSUP.Visible = True
        lblTitle(10).Visible = True
        lblTitle(10).Caption = "E.I. Code"
        cmbUIC.Visible = True
        If glbCompSerial = "S/N - 2486W" Then 'Ticket #30426 Franks 07/26/2017
            'hide this field - W.S.I.B.
        Else
            lblTitle(11).Visible = True
            txtWCB.Visible = True
        End If
        lblTitle(12).Visible = True
        cmbPenCode.Visible = True
        lblTitle(19).Visible = True
        lblTitle(19).Caption = "C.P.P."
        cmbCPP.Visible = True
        lblTitle(35).Visible = True
        txtWSIBCde.Visible = True
        Label2.Visible = True
        txtPAYFREQ.Visible = True
        clpCode(1).ShowDescription = True
        If glbCompSerial = "S/N - 2417W" Then  'Ticket #22710 - County of Perth
        Else
        medVacPPct.Visible = True
        lblTitle(8).Visible = True
        End If
        clpProv.Visible = True
        lblTitle(9).Visible = True
        medPenPct.Visible = True
        lblTitle(4).Visible = True
        lblVadim1.Visible = True
        lblVadim2.Visible = True
        clpVadim1.Visible = True
        clpVadim2.Visible = True
    End If
If glbPayWeb Then
    lblTitle(11) = "E.I. Reduced Rate"
    txtWCB.Tag = "30-E.I. Reduced Rate Y=Yes N=No"
    lblTitle(5).FontBold = True
    lblTitle(9).FontBold = True
    lblTitle(9) = "Prov. of Employment"
    lblTitle(10).FontBold = True
    lblTitle(11).FontBold = True
    lblTitle(19).FontBold = True
    lblTitle(27).FontBold = True
    lblTitle(35).FontBold = True
End If
If glbInsync Then
    lblTitle(11) = "Status Federal Tax"
    lblTitle(11).Visible = True
    cmbWCB.Tag = "10-Choose Status Federal Tax"
    If glbCompSerial <> "S/N - 2439W" Then   'OK Tire - Ticket #22128
        lblTitle(5).FontBold = True
    End If
    lblTitle(9).FontBold = True
    lblTitle(9) = "Prov. of Employment"
    If glbCompSerial <> "S/N - 2439W" Then   'OK Tire - Ticket #22128
        lblTitle(10).FontBold = True
        lblTitle(27).FontBold = True
    End If
    cmbWCB.Visible = True
    txtWCB.Visible = False
End If
If glbVadim Then
    lblTitle(11) = "E.I. Rate"
    cmbWCB.Tag = "10-Choose EI Rate"
    lblTitle(5).FontBold = True 'TD1 DOLLAR
    lblTitle(9).FontBold = True 'EMP PROV
    lblTitle(10).FontBold = True 'EI CODE
    lblTitle(27).FontBold = True 'PROV AMOUNT
    lblTitle(19).FontBold = True 'CPP
    If glbCompSerial <> "S/N - 2362W" Then 'sarnia must be removed if the integration is on
        lblTitle(11).FontBold = True 'EI RATE
        'lblTitle(35).FontBold = True 'WSIB CODE    'Hemu - for all not mandatory
        lblCalcCode.FontBold = True 'INCOME TAX
    End If
    
    'Town of Lasalle
    If glbCompSerial = "S/N - 2379W" Then
        lblTitle(11).FontBold = False 'EI RATE
    End If

    'Town of Lasalle - not for them
    If glbCompSerial <> "S/N - 2379W" Then
        cmbWCB.Visible = True
        txtWCB.Visible = False
        
        cmbWSIBCode.Visible = True
        txtWSIBCde.Visible = False
    End If
    
    If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
        txtPAYFREQ.Visible = False
    End If
    
    'Ticket #20931 - Town of Aurora
    If glbCompSerial = "S/N - 2378W" Then
        lblTitle(10).Caption = "E.I. Applicable"
        lblTitle(35).FontBold = True
    End If
    
End If

If glbCompSerial = "S/N - 2357W" And glbEmpCountry = "U.S.A." Then   'I.T. Xchange
    lblTitle(12).Visible = True
    lblTitle(12).Caption = "Federal Filing Status"
    cmbPenCode.Visible = True
    cmbPenCode.Tag = "10-Choose Federal Filing Status Code"
    Call ComPenCode
End If

If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Then    'City of Kawartha Lakes & City of Timmins
    lblTitle(12).Visible = True
    lblTitle(12).Caption = "RPP Code"
    cmbPenCode.Visible = True
    cmbPenCode.Tag = "10-Choose RPP Code"
    Call ComPenCode
End If

If glbCompSerial = "S/N - 2447W" Then    'Town of Greater Napanee 'Ticket #22905
    lblTitle(12).Visible = True
    lblTitle(12).Caption = "OMERS"
    cmbPenCode.Visible = True
    cmbPenCode.Tag = "10-Choose OMERS"
    Call ComPenCode
End If


'Ticket #18739 06/28/2010 Frank
'If glbWFC And glbCountry = "U.S.A." Then 'And fgetSection(lblEEID.Caption) = "GREN" Then
If glbWFC And glbEmpCountry = "U.S.A." Then
    frmGeneral.Visible = False
    frmlinamar.Visible = False
    fraUSA.Visible = True
    txtUIC.DataField = ""
    txtFedExemp.DataField = "ED_UIC"
    txtCPP.DataField = ""
    txtStateExemption.DataField = "ED_CPP"
    medTD3.DataField = ""
    medStateExtra.DataField = "ED_TD3"
    medTD3PC.DataField = ""
    medStateExtraPC.DataField = "ED_TD3PC"
    MedExtraTax.DataField = ""
    medFedExtra.DataField = "ED_ExtraTax"
    medExtraTaxPC.DataField = ""
    medFedExtraPC.DataField = "ED_Extrataxpc"
    txtGrossCalc.DataField = ""
    txtFedMarry.DataField = "ED_GROSSCD"
    txtWSIBCde.DataField = ""
    txtStateMarry.DataField = "ED_WCBCODE"
    txtStatusFalg3.DataField = "ED_PENSION"
    txtPension.DataField = ""
    'Vadim fields - Begin Ticket #16392
    'clpVadim11.DataField = "ED_VADIM1" 'Ticket #29828 Franks 02/14/2017 - don't use it since it is on Status/Date screen
    clpVadim1.DataField = ""
    clpVadim21.DataField = "ED_VADIM2" 'Ticket #29965 Franks 03/20/2017 need this to check logic for Fremont 'Ticket #29828 Franks 02/14/2017 - don't use it since it is on Status/Date screen
    clpVadim2.DataField = ""
    clpProvE.DataField = "ED_PROVEMP"
    clpProv.DataField = ""
    'Vadim fields - End
    If glbPlantCode = "DELR" Or glbPlantCode = "ELPA" Or glbPlantCode = "EPLM" Then
        cboDepositCode.Visible = True 'False
        cboDepositCode2.Visible = True 'False
        cboDepositCode3.Visible = True 'False
    Else
        cboDepositCode.Visible = False
        cboDepositCode2.Visible = False
        cboDepositCode3.Visible = False
    End If
    If glbEmpCountry = "U.S.A." Then 'Ticket #16616
        lblTitle(45).Caption = "State of Employment"
    Else
        lblTitle(45).Caption = "Province of Employment"
    End If
    'Ticket # Ticket #11773 - begin 'Disable these four fields
    'Ticket #12485 - wfc needs the two fields
    'lbltitle(38).Enabled = False
    'lbltitle(39).Enabled = False
    'medFedExtra.Enabled = False
    'medFedExtraPC.Enabled = False
    'lbltitle(43).Enabled = False
    'lbltitle(44).Enabled = False
    'medStateExtra.Enabled = False
    'medStateExtraPC.Enabled = False
    'Ticket # Ticket #11773 - end
    
    'Ticket #22553 Franks 09/17/2012 - begin
    clpCode(0).DataField = "ED_SUPCODE"
    clpCode(1).DataField = ""
    clpHOME.DataField = "ED_HOMEWRKCNT"
    clpHOME.TABLTitle = "Local Tax Code WI Code List"
    clpHOME.Tag = "00-Local Tax Code WI"
    'Ticket #22553 Franks 09/17/2012 - end
Else
    fraUSA.Visible = False
    txtUIC.DataField = "ED_UIC"
    txtFedExemp.DataField = ""
    txtCPP.DataField = "ED_CPP"
    txtStateExemption.DataField = ""
    medTD3.DataField = "ED_TD3"
    medStateExtra.DataField = ""
    medTD3PC.DataField = "ED_TD3PC"
    medStateExtraPC.DataField = ""
    MedExtraTax.DataField = "ED_ExtraTax"
    medFedExtra.DataField = ""
    medExtraTaxPC.DataField = "ED_Extrataxpc"
    medFedExtraPC.DataField = ""
    txtGrossCalc.DataField = "ED_GROSSCD"
    txtFedMarry.DataField = ""
    txtWSIBCde.DataField = "ED_WCBCODE"
    txtStateMarry.DataField = ""
    txtStatusFalg3.DataField = ""
    txtPension.DataField = "ED_PENSION"
    'Ticket #22553 Franks 09/18/2012 - begin
    clpCode(0).DataField = ""
    clpCode(1).DataField = "ED_SUPCODE"
    'Ticket #22553 Franks 09/18/2012 - end
    cboDepositCode.Visible = False
    cboDepositCode2.Visible = False
    cboDepositCode3.Visible = False
    'Ticket #18739  06/28/2010 Frank
    'If glbWFC And glbCountry = "CANADA" Then
    If glbWFC And glbEmpCountry = "CANADA" Then
        'Ticket #11941
        'don't have drop down list for EI and CPP for Canada user.
        'Ticket #20049 Franks 04/06/2011 - Jerry asks to not show EI, CPP and Feb Tax for Canadian employees
        lblTitle(10).Visible = False 'cpp
        lblTitle(19).Visible = False 'cpp
        lblCalcCode.Visible = False
        txtGrossCalc.Visible = False
        'Ticket #20049 - end
        cmbCPP.Visible = False
        'txtCPP.Visible = True 'Ticket #20049
        txtCPP.Top = cmbCPP.Top
        txtCPP.Left = cmbCPP.Left
        cmbUIC.Visible = False
        'txtUIC.Visible = True 'Ticket #20049
        txtUIC.Left = cmbUIC.Left
        txtUIC.Top = cmbUIC.Top
        
        'Ticket #12019 - Begin
        'Hide what the Canadian Payspecialist doesn't use
        'only keep Income Tax Applicable, CPP, EI
        lblSupervisor.Visible = False
        clpCode(1).Visible = False
        lblGarn.Visible = False
        txtGarn.Visible = False
        lblTitle(21).Visible = False
        medCHDSUP.Visible = False
        lblVadim1.Visible = False
        lblVadim2.Visible = False
        clpVadim1.Visible = False
        clpVadim2.Visible = False
        lblTitle(35).Visible = False
        cmbWSIBCode.Visible = False
        Label2.Visible = False
        cmbPayFreq.Visible = False
        lblTitle(4).Visible = False
        medPenPct.Visible = False
        txtWSIBCde.Visible = False
        txtPAYFREQ.Visible = False
        lblTitle(12).Visible = False
        cmbPenCode.Visible = False
        lblTitle(11).Visible = False
        cmbWCB.Visible = False
        txtWCB.Visible = False
        lblTitle(8).Visible = False
        medVacPPct.Visible = False
        lblCalcCode.Caption = "CPP/QPP"
        lblTitle(19).Caption = "Federal Tax"
        'Ticket #12019 - End
    End If
End If

If glbWFC Then 'Ticket #17823 Mar 22, 2010
    'From Jerry's email: They should be bolded if mandatory.
    If glbEmpCountry = "U.S.A." Then
        lblVadim21.FontBold = True
        lblTitle(45).FontBold = True
    Else
        lblVadim21.FontBold = False
        lblTitle(45).FontBold = False
    End If
    'Ticket #29828 Franks 02/14/2017 - don't use it since it is on Status/Date screen - begin
    clpVadim1.DataField = ""
    clpVadim2.DataField = ""
    lblVadim1.Visible = False
    clpVadim1.Visible = False
    lblVadim2.Visible = False
    clpVadim2.Visible = False
    'Ticket #29828 Franks 02/14/2017 - don't use it since it is on Status/Date screen - end
End If

If glbLinamar Then
    'Ticket #20188 - Begin - Increase the Bank Code for non Canadian and US employees to 5
    If glbEmpCountry <> "CANADA" And glbEmpCountry <> "U.S.A." Then
        txtBankCode.MaxLength = 5
        txtBankCode2.MaxLength = 5
        txtBankCode3.MaxLength = 5
    Else
        txtBankCode.MaxLength = 4
        txtBankCode2.MaxLength = 4
        txtBankCode3.MaxLength = 4
    End If
Else
    txtBankCode.MaxLength = 4
    txtBankCode2.MaxLength = 4
    txtBankCode3.MaxLength = 4
End If

If glbCompSerial = "S/N - 2382W" Then 'Ticket #18090 Samuel
    Call SamuelScreenSetup
End If

'Ticket #28786 - Goodmans
If glbCompSerial = "S/N - 2290W" Then
    lblTitle(9).FontBold = True 'Province of Employment
    lblTitle(10).FontBold = True 'E.I.
    lblTitle(19).FontBold = True 'CPP
    lblTitle(11).Caption = "EI Pref."
End If

If glbCompSerial = "S/N - 2487W" Then 'Ticket #30217 Franks 06/13/2017 City of Kenora
    Call ScreenSetupKenora
End If

End Sub

Private Function fgetSection(xID) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As String
    
    If glbtermopen Then
        strSQL = "SELECT ED_SECTION FROM TERM_HREMP WHERE TERM_SEQ =" & glbTERM_ID
        rs.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
    Else
        strSQL = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    End If
    
    If rs.EOF = False Then
        If Not IsNull(rs("ED_SECTION")) Then
            retVal = rs("ED_SECTION")
        Else
            retVal = ""
        End If
    Else
        retVal = ""
    End If
    rs.Close
    Set rs = Nothing
    
    fgetSection = retVal

End Function

Private Sub Pass_TermBank_Changes_Vadim()
Dim xOVadim1, xOVadim2 As Double
Dim UpdateAudit As Boolean
Dim Banks As New Collection

If isChanged_Bank(Banks, OBANK, txtBankCode) Then UpdateAudit = True
If isChanged_Bank(Banks, OBRANCH, txtBranchCode) Then UpdateAudit = True
If isChanged_Bank(Banks, OACCOUNT, txtAccount) Then UpdateAudit = True
If isChanged_Bank(Banks, OAMTDEPOSIT, medAmountDeposit, True) Then UpdateAudit = True
If isChanged_Bank(Banks, OPCDEPOSIT, medPCDeposit, True) Then UpdateAudit = True
If isChanged_Bank(Banks, ODEPCODE, txtDepositCode) Then UpdateAudit = True

If isChanged_Bank(Banks, OBANK2, txtBankCode2) Then UpdateAudit = True
If isChanged_Bank(Banks, OBRANCH2, txtBranchCode2) Then UpdateAudit = True
If isChanged_Bank(Banks, OACCOUNT2, txtAccount2) Then UpdateAudit = True
If isChanged_Bank(Banks, OAMTDEPOSIT2, medAmountDeposit2, True) Then UpdateAudit = True
If isChanged_Bank(Banks, OPCDEPOSIT2, medPCDeposit2, True) Then UpdateAudit = True
If isChanged_Bank(Banks, ODEPCODE2, txtDepositCode2) Then UpdateAudit = True

If isChanged_Bank(Banks, OBANK3, txtBankCode3) Then UpdateAudit = True
If isChanged_Bank(Banks, OBRANCH3, txtBranchCode3) Then UpdateAudit = True
If isChanged_Bank(Banks, OACCOUNT3, txtAccount3) Then UpdateAudit = True
If isChanged_Bank(Banks, OAMTDEPOSIT3, medAmountDeposit3, True) Then UpdateAudit = True
If isChanged_Bank(Banks, OPCDEPOSIT3, medPCDeposit3, True) Then UpdateAudit = True
If isChanged_Bank(Banks, ODEPCODE3, txtDepositCode3) Then UpdateAudit = True

If UpdateAudit Then Call Passing_Bank_Changes(Banks, glbTERM_ID, rsDATA("ED_Payroll_ID"))

Dim HRChanges As New Collection
If isChanged_Field(HRChanges, OVACPC, medVacPPct) Then UpdateAudit = True
If isChanged_Field(HRChanges, OWCB, txtWCB) Then UpdateAudit = True
If isChanged_Field(HRChanges, OPENSION, txtPension) Then UpdateAudit = True
If isChanged_Field(HRChanges, OWSIBCDE, txtWSIBCde) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1CODE, txtTD1Code) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1DOL, medTD1Amnt, True) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD3, medTD3, True) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTD1, lblTD1) Then UpdateAudit = True
If isChanged_Field(HRChanges, OSUPCODE, clpCode(1)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODDI, lblDirectDeposit) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvEmp, clpProv) Then UpdateAudit = True
If glbWFC And fraUSA.Visible Then
    If isChanged_Field(HRChanges, OUIC, txtFedExemp) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OCPP, txtStateExemption) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OGROSCALC, txtFedMarry) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTax, medFedExtra, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTaxPC, medFedExtraPC, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oStateExtraTax, medStateExtra, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oStateExtraTaxPC, medStateExtraPC, True) Then UpdateAudit = True
Else
    If isChanged_Field(HRChanges, OUIC, txtUIC) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OCPP, txtCPP) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OGROSCALC, txtGrossCalc) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTax, MedExtraTax, True) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oExtraTaxPC, medExtraTaxPC, True) Then UpdateAudit = True
End If
If isChanged_Field(HRChanges, OGARN, txtGarn) Then UpdateAudit = True
If isChanged_Field(HRChanges, oFedTax, txtFedTax) Then UpdateAudit = True
If isChanged_Field(HRChanges, oExtAmt, txtExtAmt) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvForm, lblProvForm) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvAmt, medProvAmt, True) Then UpdateAudit = True
If isChanged_Field(HRChanges, oProvCode, txtProvCode) Then UpdateAudit = True
If isChanged_Field(HRChanges, OExtrAnn, txtExtrAnn) Then UpdateAudit = True
If isChanged_Field(HRChanges, OQTBTORRSP, lblQTBTORRSP) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTADDR, txtOUTAddr) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTCITY, txtOUTCity) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTPROV, clpOUTProv) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTCOUNTRY, comOUTCountry) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTPCODE, medOUTPCode) Then UpdateAudit = True
If isChanged_Field(HRChanges, OOUTADDRT4, chkOUTADDRT4) Then UpdateAudit = True

If isChanged_Field(HRChanges, OTRANSITABA, txtTransitABA) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTRANSITABA2, txtTransitABA2) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTRANSITABA3, txtTransitABA3) Then UpdateAudit = True
If isChanged_Field(HRChanges, OPenPct, medPenPct) Then UpdateAudit = True

If glbCompSerial = "S/N - 2276W" Then    'City of Niagara Falls
    'Federal Alimony Child Support
    If isChanged_Field(HRChanges, oFedAliChd, medCHDSUP) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oPAYFREQ, txtPAYFREQ) Then UpdateAudit = True
End If

'Vadim Field 1
'For City of Timmins or City of Niagara Falls
If (glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W") And clpVadim1 <> "" Then
    medVadim1.DataField = "ED_VADIM1"
    medVadim1 = (Val(clpVadim1) / 100)
    xOVadim1 = (Val(oVadim1) / 100)
    If isChanged_Field(HRChanges, xOVadim1, medVadim1, True) Then UpdateAudit = True
    medVadim1.DataField = ""
Else
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
        If isChanged_Field(HRChanges, oVadim1, clpVadim1, True) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChanges, oVadim1, clpVadim1) Then UpdateAudit = True
    End If
End If

'Vadim Field 2
'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then 'Or glbCompSerial = "S/N - 2375W" Then
    If clpVadim2 <> "" Then
        medVadim2.DataField = "ED_VADIM2"
        medVadim2 = (Val(clpVadim2) / 100)
        xOVadim2 = (Val(OVadim2) / 100)
        If isChanged_Field(HRChanges, xOVadim2, medVadim2, True) Then UpdateAudit = True
        medVadim2.DataField = ""
    Else
        xOVadim2 = (Val(OVadim2) / 100)
        If isChanged_Field(HRChanges, xOVadim2, clpVadim2, True) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, OVadim2, clpVadim2) Then UpdateAudit = True
End If

Call Passing_Changes(HRChanges, Banking, "M", Date, glbTERM_ID, rsDATA("ED_Payroll_ID"))

End Sub

Private Sub ScreenSetupKenora()
'Hide some fields
lblTitle(6).Visible = False: txtTD1Code.Visible = False
lblTitle(20).Visible = False: medTD3PC.Visible = False
lblTitle(23).Visible = False: txtFedTax.Visible = False
lblTitle(22).Visible = False: txtExtAmt.Visible = False
lblTitle(26).Visible = False: txtProvCode.Visible = False
lblTitle(24).Visible = False: medExtraTaxPC.Visible = False

lblTitle(9).Visible = False: clpProv.Visible = False
lblSupervisor.Visible = False: clpCode(1).Visible = False
lblCalcCode.Visible = False: txtGrossCalc.Visible = False
lblGarn.Visible = False: txtGarn.Visible = False
lblTitle(21).Visible = False: medCHDSUP.Visible = False
lblVadim1.Visible = False: clpVadim1.Visible = False
lblVadim2.Visible = False: clpVadim2.Visible = False

lblTitle(10).Visible = False: cmbUIC.Visible = False
lblTitle(11).Visible = False: cmbWCB.Visible = False: txtWCB.Visible = False
lblTitle(12).Visible = False: cmbPenCode.Visible = False
lblTitle(19).Visible = False: cmbCPP.Visible = False
'lbltitle(35).Visible = False: cmbWSIBCode.Visible = False
Label2.Visible = False: cmbPayFreq.Visible = False: txtPAYFREQ.Visible = False
lblTitle(4).Visible = False: medPenPct.Visible = False

End Sub

Private Sub SamuelScreenSetup()
    lblTitle(5).FontBold = True 'TD1 Amount
    lblTitle(19).FontBold = True 'C.P.P
    lblTitle(27).FontBold = True 'Prov. Amount
    
    'Ticket #18417 - begin Frank 04/27/2010
    lblTitle(35).Caption = "QPIP"
    cmbWSIBCode.Visible = True
    cmbWSIBCode.Clear
    cmbWSIBCode.AddItem "0 - Subject to QPIP"
    cmbWSIBCode.AddItem "X - Exempt"
    cmbWSIBCode.AddItem ""
    'Ticket #18417 - end
    
    'Ticket #19938 Franks 05/26/2011 - begin
    lblPPercFixed.Visible = True
    chkPenFixed.Visible = True
    chkPenFixed.DataField = "ED_PENPCTFIXED"
    'Ticket #19938 Franks 05/26/2011 - end
    
    'Ticket #20600 Franks 09/01/2011
    'Vadim Field 1
    lblVadim1.Visible = False
    clpVadim1.Visible = False
    'Supervisor Code
    lblSupervisor.Visible = False
    clpCode(1).Visible = False
    
End Sub


