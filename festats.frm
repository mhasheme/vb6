VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEESTATS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "EMP Status/Dates"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   1425
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10515
   ScaleWidth      =   14535
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   3645
      LargeChange     =   315
      Left            =   11400
      Max             =   100
      SmallChange     =   315
      TabIndex        =   70
      Top             =   4200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame TopFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   120
      TabIndex        =   71
      Top             =   480
      Width           =   14295
      Begin VB.CommandButton cmdLOAComments 
         Appearance      =   0  'Flat
         Caption         =   "LOA Comments"
         Height          =   330
         Left            =   11640
         TabIndex        =   4
         Tag             =   "Terminate the Employee Selected"
         Top             =   97
         Width           =   1500
      End
      Begin VB.CommandButton cmdImport2 
         Caption         =   "Import"
         Height          =   330
         Left            =   10635
         TabIndex        =   3
         Top             =   97
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox comEmpType 
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
         ItemData        =   "festats.frx":0000
         Left            =   2550
         List            =   "festats.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "10-Type of Employee "
         Top             =   430
         Width           =   2655
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   255
         Index           =   4
         Left            =   2235
         TabIndex        =   110
         Tag             =   "00-Section"
         Top             =   435
         Visible         =   0   'False
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
         Object.Height          =   255
      End
      Begin VB.CommandButton cmdEmailImpFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   13035
         TabIndex        =   16
         Tag             =   "Select File to Import"
         Top             =   1742
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdEmailImp 
         Appearance      =   0  'Flat
         Caption         =   "Import"
         Height          =   280
         Left            =   13440
         TabIndex        =   17
         Tag             =   "Import the File"
         Top             =   1742
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtUserText2 
         Appearance      =   0  'Flat
         DataField       =   "ED_USER_TEXT2"
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
         Left            =   8385
         MaxLength       =   20
         TabIndex        =   22
         Tag             =   "00-User Text 2"
         Top             =   2400
         Width           =   1620
      End
      Begin VB.ComboBox comUserText2 
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
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   191
         Tag             =   "00-User Text 2"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.ComboBox comUserText1 
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
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "00-User Text 1"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditNGSSub 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   11640
         TabIndex        =   183
         Tag             =   "Edit Transaction Date"
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtUserNum2 
         Appearance      =   0  'Flat
         DataField       =   "ED_USER_NUM2"
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
         Left            =   8385
         MaxLength       =   20
         TabIndex        =   24
         Tag             =   "00-User Number 2"
         Top             =   2730
         Width           =   1620
      End
      Begin VB.CommandButton cmdEditUserNum1 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1900
         TabIndex        =   177
         Tag             =   "Edit Transaction Date"
         Top             =   2730
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtUserNum1 
         Appearance      =   0  'Flat
         DataField       =   "ED_USER_NUM1"
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
         Left            =   2550
         MaxLength       =   20
         TabIndex        =   23
         Tag             =   "00-User Number 1"
         Top             =   2730
         Width           =   1620
      End
      Begin VB.CommandButton cmdEditUserText2 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7695
         TabIndex        =   105
         Tag             =   "Edit Transaction Date"
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdEditUserText1 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1900
         TabIndex        =   104
         Tag             =   "Edit Transaction Date"
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtUserText1 
         Appearance      =   0  'Flat
         DataField       =   "ED_USER_TEXT1"
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
         Height          =   280
         Left            =   2550
         MaxLength       =   20
         TabIndex        =   20
         Tag             =   "00-User Text 1"
         Top             =   2400
         Width           =   1620
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   330
         Left            =   8385
         TabIndex        =   102
         Top             =   3135
         Visible         =   0   'False
         Width           =   855
      End
      Begin INFOHR_Controls.CodeLookup clpSalDist 
         DataField       =   "ED_SALDIST"
         Height          =   285
         Left            =   8070
         TabIndex        =   14
         Top             =   1740
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "ED_SFDATE"
         DataSource      =   " "
         Height          =   285
         Index           =   15
         Left            =   6000
         TabIndex        =   1
         Tag             =   "40-Status From Date"
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_HIRECODE"
         DataSource      =   " "
         Height          =   285
         Index           =   6
         Left            =   2240
         TabIndex        =   18
         Tag             =   "Hire Code"
         Top             =   2070
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDHC"
      End
      Begin VB.CheckBox chkLeave 
         Caption         =   "Leave"
         Height          =   195
         Left            =   5745
         TabIndex        =   97
         Top             =   3240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtIPHONE 
         Appearance      =   0  'Flat
         DataField       =   "ED_INTEL"
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
         Left            =   2550
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "00-Internal Telephone Extension "
         Top             =   1410
         Width           =   1305
      End
      Begin VB.TextBox lblDob 
         Appearance      =   0  'Flat
         DataField       =   "ED_DOB"
         DataSource      =   " "
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
         Left            =   11760
         MaxLength       =   25
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "lblDob"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LDATE"
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
         Left            =   12240
         MaxLength       =   25
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "Ldate"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LTIME"
         DataSource      =   " "
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
         Left            =   12720
         MaxLength       =   25
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "LTime"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtEmpType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_EMPTYPE"
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
         Left            =   7080
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4020
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtfDate 
         Appearance      =   0  'Flat
         DataField       =   "ED_EFDATE"
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
         Left            =   8040
         MaxLength       =   11
         TabIndex        =   80
         Tag             =   "41-Original Date Hired "
         Text            =   " "
         Top             =   3900
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtTDate 
         Appearance      =   0  'Flat
         DataField       =   "ED_ETDATE"
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
         Left            =   8040
         MaxLength       =   11
         TabIndex        =   79
         Tag             =   "41-Original Date Hired "
         Top             =   4260
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtVACT 
         Appearance      =   0  'Flat
         DataField       =   "ED_VACT"
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
         Left            =   8010
         MaxLength       =   10
         TabIndex        =   78
         Tag             =   "41-Original Date Hired "
         Top             =   4590
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtSICKT 
         Appearance      =   0  'Flat
         DataField       =   "ED_SICKT"
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
         Left            =   9270
         MaxLength       =   10
         TabIndex        =   77
         Tag             =   "41-Original Date Hired "
         Top             =   4590
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtFDateS 
         Appearance      =   0  'Flat
         DataField       =   "ED_EFDATES"
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
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   76
         Tag             =   "41-Original Date Hired "
         Top             =   3900
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtTDateS 
         Appearance      =   0  'Flat
         DataField       =   "ED_ETDATES"
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
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   75
         Tag             =   "41-Original Date Hired "
         Top             =   4260
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LUSER"
         DataSource      =   " "
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
         Left            =   13140
         MaxLength       =   25
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   "LUser"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         DataField       =   "ED_EMAIL"
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
         Left            =   2550
         MaxLength       =   60
         TabIndex        =   13
         Tag             =   "00-Email Address"
         Top             =   1740
         Width           =   3950
      End
      Begin VB.CheckBox chkSpouse 
         DataField       =   "ED_WITHSPOUSE"
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
         Height          =   255
         Left            =   2600
         TabIndex        =   28
         Tag             =   "Spouse Works at Linamar"
         Top             =   3110
         Width           =   615
      End
      Begin VB.TextBox txtPVAC 
         Appearance      =   0  'Flat
         DataField       =   "ED_PVAC"
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   73
         Tag             =   "41-Original Date Hired "
         Top             =   4860
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txtVAC 
         Appearance      =   0  'Flat
         DataField       =   "ED_VAC"
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
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   72
         Tag             =   "41-Original Date Hired "
         Top             =   4860
         Visible         =   0   'False
         Width           =   1230
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         DataField       =   "ED_PT"
         DataSource      =   " "
         Height          =   285
         Left            =   2240
         TabIndex        =   7
         Tag             =   "00-Category Codes"
         Top             =   750
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_EMP"
         Height          =   285
         Index           =   1
         Left            =   2240
         TabIndex        =   0
         Tag             =   "00-Enter Status Code"
         Top             =   120
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ORG"
         DataSource      =   " "
         Height          =   285
         Index           =   2
         Left            =   2240
         TabIndex        =   9
         Tag             =   "00-Enter Union Code"
         Top             =   1080
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpBGroup 
         DataField       =   "ED_BENEFIT_GROUP"
         Height          =   285
         Left            =   8070
         TabIndex        =   12
         Tag             =   "01-Benefit - Group Code"
         Top             =   1410
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
         SecurityMaintainable=   0
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   17
         Left            =   8070
         TabIndex        =   109
         Tag             =   "40-Date"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1045
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   288
         Index           =   7
         Left            =   2240
         TabIndex        =   6
         Top             =   433
         Visible         =   0   'False
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "ED_STDATE"
         DataSource      =   " "
         Height          =   285
         Index           =   16
         Left            =   8070
         TabIndex        =   2
         Tag             =   "40-Status To Date"
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "ED_PTEDATE"
         DataSource      =   " "
         Height          =   285
         Index           =   34
         Left            =   8070
         TabIndex        =   8
         Tag             =   "40-Category Effective Date"
         Top             =   750
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpVadim1 
         Height          =   285
         Left            =   12360
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim2 
         Height          =   285
         Left            =   12360
         TabIndex        =   26
         Top             =   2730
         Visible         =   0   'False
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV2"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   12435
         TabIndex        =   25
         Tag             =   "00-Supervisory Code for cheque sorting "
         Top             =   3240
         Visible         =   0   'False
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSP"
      End
      Begin MSMask.MaskEdBox medVacPPct 
         Height          =   285
         Left            =   12720
         TabIndex        =   15
         Tag             =   "10-Enter Vacation Pay Percentage "
         Top             =   1320
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_LOC"
         Height          =   285
         Index           =   0
         Left            =   8070
         TabIndex        =   19
         Tag             =   "00-Location - Code"
         Top             =   2070
         Visible         =   0   'False
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "ED_ORGEDATE"
         DataSource      =   " "
         Height          =   285
         Index           =   36
         Left            =   8070
         TabIndex        =   10
         Tag             =   "40-Union Effective Date"
         Top             =   1080
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lblUnionEDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Union Effective"
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
         Left            =   6840
         TabIndex        =   200
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
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
         Left            =   12960
         TabIndex        =   199
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image imgNoSec2 
         Height          =   240
         Left            =   10215
         Picture         =   "festats.frx":0004
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LOA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   9720
         TabIndex        =   198
         Top             =   120
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image imgSec2 
         Height          =   240
         Left            =   10215
         Picture         =   "festats.frx":014E
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblSection 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6675
         TabIndex        =   197
         Top             =   2115
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label txtFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8535
         TabIndex        =   192
         Top             =   1785
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Image imgHelp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   12720
         Picture         =   "festats.frx":0298
         Stretch         =   -1  'True
         Top             =   1755
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblSupervisor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Code"
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
         Left            =   9960
         TabIndex        =   188
         Top             =   3240
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblWFCMsg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "msg"
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
         Left            =   12600
         TabIndex        =   187
         Top             =   3000
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblVadim11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
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
         Left            =   9960
         TabIndex        =   182
         Top             =   2400
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblVadim21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 2"
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
         Left            =   9960
         TabIndex        =   181
         Top             =   2730
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblUserNum2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Number 2"
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
         Left            =   6000
         TabIndex        =   180
         Top             =   2775
         Width           =   1545
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   7740
         TabIndex        =   176
         Top             =   165
         Width           =   195
      End
      Begin VB.Label lblPTEDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Effective"
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
         Left            =   6630
         TabIndex        =   175
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblUserNum1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Number 1"
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
         Left            =   0
         TabIndex        =   107
         Top             =   2775
         Width           =   1905
      End
      Begin VB.Label lblUserText2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Text 2"
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
         Left            =   5760
         TabIndex        =   106
         Top             =   2445
         Width           =   1785
      End
      Begin VB.Image imgNoSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8070
         Picture         =   "festats.frx":06DA
         Top             =   3180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8070
         Picture         =   "festats.frx":0824
         Top             =   3180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         Caption         =   "Resume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6960
         TabIndex        =   101
         Top             =   3180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSalDist 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Distribution"
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
         Left            =   6675
         TabIndex        =   100
         Top             =   1785
         Width           =   1260
      End
      Begin VB.Label lblBen 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6960
         TabIndex        =   98
         Top             =   1455
         Width           =   975
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   96
         Top             =   795
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   5520
         TabIndex        =   95
         Top             =   165
         Width           =   345
      End
      Begin VB.Label lblEEStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   94
         Top             =   165
         Width           =   2205
      End
      Begin VB.Label lblEEType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Type"
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
         Left            =   0
         TabIndex        =   93
         Top             =   490
         Width           =   1965
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   0
         TabIndex        =   92
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblIPhone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Phone Extension"
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
         Left            =   0
         TabIndex        =   91
         Top             =   1455
         Width           =   2130
      End
      Begin VB.Label lblUserText1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Text 1"
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
         Left            =   0
         TabIndex        =   90
         Top             =   2443
         Width           =   1905
      End
      Begin VB.Label lblDateS 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EntoutS"
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
         Left            =   4680
         TabIndex        =   89
         Top             =   3270
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblDATE 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Entout"
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
         Left            =   4050
         TabIndex        =   88
         Top             =   3270
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblEmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   0
         TabIndex        =   87
         Top             =   1785
         Width           =   1455
      End
      Begin VB.Label lblSpouse 
         Caption         =   "Spouse Works at Linamar"
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
         Left            =   0
         TabIndex        =   86
         Top             =   3110
         Width           =   2055
      End
      Begin VB.Label lblHireCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Code"
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
         Left            =   0
         TabIndex        =   85
         Top             =   2115
         Width           =   705
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
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
         Index           =   20
         Left            =   6600
         TabIndex        =   99
         Top             =   2115
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   68
      Top             =   9960
      Width           =   14535
      _Version        =   65536
      _ExtentX        =   25638
      _ExtentY        =   979
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
      Begin VB.HScrollBar scrHScroll 
         Height          =   300
         LargeChange     =   25
         Left            =   0
         Max             =   50
         SmallChange     =   4
         TabIndex        =   103
         Top             =   0
         Width           =   11175
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   14535
      _Version        =   65536
      _ExtentX        =   25638
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
      Begin VB.TextBox txtTEmpNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   190
         Text            =   "txtTEmpNo"
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtTEmpNames 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   189
         Text            =   "txtTEmpNames"
         Top             =   120
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDemo 
         Appearance      =   0  'Flat
         Caption         =   "Comments"
         Height          =   330
         Left            =   7920
         TabIndex        =   184
         Tag             =   "Terminate the Employee Selected"
         Top             =   80
         Width           =   1500
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_SURNAME"
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
         Left            =   6960
         MaxLength       =   25
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_FNAME"
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
         Left            =   7440
         MaxLength       =   25
         TabIndex        =   63
         Text            =   "Text6"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCommDesc 
         Caption         =   "Comments Entered"
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
         Height          =   255
         Left            =   9480
         TabIndex        =   186
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
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
         Left            =   7320
         TabIndex        =   108
         Top             =   120
         Width           =   75
      End
      Begin VB.Label lblEENUM 
         AutoSize        =   -1  'True
         Caption         =   "lblEENUM"
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
         TabIndex        =   69
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   160
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3600
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   840
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
         Left            =   2760
         TabIndex        =   65
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame ScrFrame 
      BorderStyle     =   0  'None
      Height          =   5450
      Left            =   120
      TabIndex        =   111
      Top             =   4080
      Width           =   11295
      Begin VB.Frame fraDateEmp 
         Height          =   735
         Left            =   0
         TabIndex        =   121
         Top             =   960
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox txtExpYear 
            Appearance      =   0  'Flat
            DataField       =   "ED_EXPYEAR"
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
            Left            =   1720
            MaxLength       =   4
            TabIndex        =   31
            Tag             =   "00-Experience Year"
            Top             =   720
            Width           =   810
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_DIVEDATE"
            DataSource      =   " "
            Height          =   285
            Index           =   14
            Left            =   4890
            TabIndex        =   40
            Tag             =   "40-Division Start Date"
            Top             =   1800
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_DEPTEDATE"
            DataSource      =   " "
            Height          =   285
            Index           =   13
            Left            =   4890
            TabIndex        =   39
            Tag             =   "40-Department Start Date"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_FMLA"
            DataSource      =   " "
            Height          =   285
            Index           =   12
            Left            =   1410
            TabIndex        =   34
            Tag             =   "40-FMLA Date"
            Top             =   1800
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1060
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_DOH"
            DataSource      =   " "
            Height          =   285
            Index           =   7
            Left            =   1410
            TabIndex        =   29
            Tag             =   "41-Original Hire Date "
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1060
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_SENDTE"
            DataSource      =   " "
            Height          =   285
            Index           =   6
            Left            =   1410
            TabIndex        =   30
            Tag             =   "40-Seniority Date"
            Top             =   360
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1060
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_LTHIRE"
            DataSource      =   " "
            Height          =   285
            Index           =   5
            Left            =   1410
            TabIndex        =   32
            Tag             =   "40-Last Hire Date"
            Top             =   1080
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1060
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_UNION"
            DataSource      =   " "
            Height          =   285
            Index           =   4
            Left            =   1410
            TabIndex        =   33
            Tag             =   "40-Union Date"
            Top             =   1440
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1060
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_USRDAT1"
            DataSource      =   " "
            Height          =   285
            Index           =   3
            Left            =   4890
            TabIndex        =   38
            Tag             =   "40-User Defined"
            Top             =   1080
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_OMERS"
            DataSource      =   " "
            Height          =   285
            Index           =   2
            Left            =   4890
            TabIndex        =   37
            Tag             =   "40-OMERS Date"
            Top             =   720
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_LDAY"
            DataSource      =   " "
            Height          =   285
            Index           =   1
            Left            =   4890
            TabIndex        =   36
            Tag             =   "40-Last Day"
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_FDAY"
            DataSource      =   " "
            Height          =   285
            Index           =   0
            Left            =   4890
            TabIndex        =   35
            Tag             =   "40-First Day"
            Top             =   0
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin VB.Label lblOHire 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Original Hire"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   135
            Top             =   45
            Width           =   1665
         End
         Begin VB.Label lblSen 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Seniority"
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
            Left            =   0
            TabIndex        =   134
            Top             =   405
            Width           =   1545
         End
         Begin VB.Label lblLHire 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Hire"
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
            Left            =   0
            TabIndex        =   133
            Top             =   1125
            Width           =   1590
         End
         Begin VB.Label lblUDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Union Date"
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
            Left            =   0
            TabIndex        =   132
            Top             =   1485
            Width           =   1650
         End
         Begin VB.Label lblFDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "First Day"
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
            Left            =   3930
            TabIndex        =   131
            Top             =   60
            Width           =   855
         End
         Begin VB.Label lblLDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Day"
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
            Left            =   3930
            TabIndex        =   130
            Top             =   420
            Width           =   855
         End
         Begin VB.Label lblODate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OMERS Date"
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
            Left            =   3810
            TabIndex        =   129
            Top             =   780
            Width           =   975
         End
         Begin VB.Label lblUDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "User Defined"
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
            Left            =   3690
            TabIndex        =   128
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "FMLA Date"
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
            Left            =   0
            TabIndex        =   127
            Top             =   1845
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Experience Year"
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
            Left            =   0
            TabIndex        =   126
            Top             =   765
            Width           =   1530
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "Years"
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
            Index           =   0
            Left            =   2820
            TabIndex        =   125
            Top             =   390
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "Years"
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
            Index           =   1
            Left            =   2820
            TabIndex        =   124
            Top             =   30
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblDeptStart 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Depart. Start Date"
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
            Left            =   3450
            TabIndex        =   123
            Top             =   1500
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDivStart 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Division Start Date"
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
            Left            =   3450
            TabIndex        =   122
            Top             =   1860
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.TextBox txtEmpComm 
         Appearance      =   0  'Flat
         DataField       =   "ER_COMMENT"
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
         Height          =   735
         Left            =   10200
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   185
         Tag             =   "00-Comments - free form"
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdImport1 
         Caption         =   "Import"
         Height          =   330
         Left            =   8480
         TabIndex        =   178
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame fraDatePension 
         Height          =   735
         Left            =   1680
         TabIndex        =   112
         Top             =   1680
         Visible         =   0   'False
         Width           =   5535
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_LATESTR"
            DataSource      =   " "
            Height          =   285
            Index           =   11
            Left            =   1650
            TabIndex        =   45
            Tag             =   "40-Latest Retirement"
            Top             =   1035
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_NORMALR"
            DataSource      =   " "
            Height          =   285
            Index           =   10
            Left            =   1650
            TabIndex        =   44
            Tag             =   "40-Normal Retirement"
            Top             =   690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_EARLYR"
            DataSource      =   " "
            Height          =   285
            Index           =   9
            Left            =   1650
            TabIndex        =   43
            Tag             =   "40-Earliest Retirement"
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ED_ELIGIBLE"
            DataSource      =   " "
            Height          =   285
            Index           =   8
            Left            =   1650
            TabIndex        =   42
            Tag             =   "40-Eligibility"
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE3"
            DataSource      =   " "
            Height          =   285
            Index           =   20
            Left            =   6450
            TabIndex        =   48
            Tag             =   "40-Pension Date 3"
            Top             =   690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE2"
            DataSource      =   " "
            Height          =   285
            Index           =   19
            Left            =   6450
            TabIndex        =   47
            Tag             =   "40-Pension Date 2"
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE1"
            DataSource      =   " "
            Height          =   285
            Index           =   18
            Left            =   6450
            TabIndex        =   46
            Tag             =   "40-Pension Date 1"
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE6"
            DataSource      =   " "
            Height          =   285
            Index           =   23
            Left            =   6450
            TabIndex        =   51
            Tag             =   "40-Pension Date 6"
            Top             =   1725
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE5"
            DataSource      =   " "
            Height          =   285
            Index           =   22
            Left            =   6450
            TabIndex        =   50
            Tag             =   "40-Pension Date 5"
            Top             =   1380
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_PENSIONDATE4"
            DataSource      =   " "
            Height          =   285
            Index           =   21
            Left            =   6450
            TabIndex        =   49
            Tag             =   "40-Pension Date 4"
            Top             =   1035
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin Threed.SSCheck chkElig 
            Height          =   255
            Left            =   0
            TabIndex        =   41
            Tag             =   "Click to Select Rehire"
            Top             =   0
            Visible         =   0   'False
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Not Eligible                   "
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
            Alignment       =   1
            Value           =   -1  'True
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   9
            Left            =   8040
            TabIndex        =   174
            Top             =   1755
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   8
            Left            =   8040
            TabIndex        =   173
            Top             =   1410
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   7
            Left            =   8040
            TabIndex        =   172
            Top             =   1080
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   6
            Left            =   8040
            TabIndex        =   171
            Top             =   675
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   5
            Left            =   8040
            TabIndex        =   170
            Top             =   330
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   4
            Left            =   8040
            TabIndex        =   169
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 6"
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
            Left            =   4560
            TabIndex        =   158
            Top             =   1725
            Width           =   1935
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 5"
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
            Left            =   4560
            TabIndex        =   157
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 4"
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
            Left            =   4560
            TabIndex        =   156
            Top             =   1035
            Width           =   1935
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 3"
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
            Left            =   4560
            TabIndex        =   155
            Top             =   690
            Width           =   1935
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 2"
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
            Left            =   4560
            TabIndex        =   154
            Top             =   345
            Width           =   1935
         End
         Begin VB.Label lblPenDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 1"
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
            Left            =   4560
            TabIndex        =   153
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   3
            Left            =   3165
            TabIndex        =   120
            Top             =   1065
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   0
            Left            =   3165
            TabIndex        =   119
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   1
            Left            =   3165
            TabIndex        =   118
            Top             =   375
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Index           =   2
            Left            =   3165
            TabIndex        =   117
            Top             =   720
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblEarlR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Earliest Retirement"
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
            Left            =   0
            TabIndex        =   116
            Top             =   345
            Width           =   1560
         End
         Begin VB.Label lblNorR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Normal Retirement"
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
            Left            =   0
            TabIndex        =   115
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label lblLateR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Retirement"
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
            Left            =   0
            TabIndex        =   114
            Top             =   1035
            Width           =   1605
         End
         Begin VB.Label lblElig 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Eligibility"
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
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   1785
         End
      End
      Begin VB.Frame fraDateOther 
         Height          =   735
         Left            =   1680
         TabIndex        =   152
         Top             =   2400
         Visible         =   0   'False
         Width           =   8055
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE3"
            DataSource      =   " "
            Height          =   285
            Index           =   26
            Left            =   1890
            TabIndex        =   54
            Tag             =   "40-Other Date 3"
            Top             =   690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE2"
            DataSource      =   " "
            Height          =   285
            Index           =   25
            Left            =   1890
            TabIndex        =   53
            Tag             =   "40-Other Date 2"
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE1"
            DataSource      =   " "
            Height          =   285
            Index           =   24
            Left            =   1890
            TabIndex        =   52
            Tag             =   "40-Other Date 1"
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE5"
            DataSource      =   " "
            Height          =   285
            Index           =   28
            Left            =   1890
            TabIndex        =   56
            Tag             =   "40-Other Date 5"
            Top             =   1380
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE4"
            DataSource      =   " "
            Height          =   285
            Index           =   27
            Left            =   1890
            TabIndex        =   55
            Tag             =   "40-Other Date 4"
            Top             =   1035
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE8"
            DataSource      =   " "
            Height          =   285
            Index           =   31
            Left            =   5610
            TabIndex        =   59
            Tag             =   "40-Other Date 8"
            Top             =   690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE7"
            DataSource      =   " "
            Height          =   285
            Index           =   30
            Left            =   5610
            TabIndex        =   58
            Tag             =   "40-Other Date 7"
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE6"
            DataSource      =   " "
            Height          =   285
            Index           =   29
            Left            =   5610
            TabIndex        =   57
            Tag             =   "40-Other Date 6"
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE10"
            DataSource      =   " "
            Height          =   285
            Index           =   33
            Left            =   5610
            TabIndex        =   61
            Tag             =   "40-Other Date 10"
            Top             =   1380
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "ER_OTHERDATE9"
            DataSource      =   " "
            Height          =   285
            Index           =   32
            Left            =   5610
            TabIndex        =   60
            Tag             =   "40-Other Date 9"
            Top             =   1035
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1045
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 6"
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
            TabIndex        =   168
            Top             =   0
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 7"
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
            TabIndex        =   167
            Top             =   345
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 8"
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
            TabIndex        =   166
            Top             =   690
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 9"
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
            TabIndex        =   165
            Top             =   1035
            Width           =   1755
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 10"
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
            TabIndex        =   164
            Top             =   1380
            Width           =   1725
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 1"
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
            Left            =   0
            TabIndex        =   163
            Top             =   0
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 2"
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
            Left            =   0
            TabIndex        =   162
            Top             =   345
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 3"
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
            Left            =   0
            TabIndex        =   161
            Top             =   690
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 4"
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
            Left            =   0
            TabIndex        =   160
            Top             =   1035
            Width           =   1875
         End
         Begin VB.Label lbOtherDate 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 5"
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
            Left            =   0
            TabIndex        =   159
            Top             =   1380
            Width           =   1875
         End
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
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
         Height          =   975
         Left            =   1400
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   141
         Tag             =   "00-Comments - free form"
         Top             =   3900
         Visible         =   0   'False
         Width           =   8200
      End
      Begin VB.TextBox txtRPP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   10680
         MaxLength       =   2
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtVadim1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   450
      End
      Begin INFOHR_Controls.DateLookup dlpRehired 
         Height          =   285
         Left            =   7830
         TabIndex        =   138
         Tag             =   "41-Date Rehired"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   1400
         TabIndex        =   139
         Tag             =   "Date Terminated"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataSource      =   " "
         Height          =   285
         Index           =   5
         Left            =   3810
         TabIndex        =   140
         Tag             =   "Termination Code - Code "
         Top             =   3180
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin Threed.SSCheck chkRehire 
         Height          =   255
         Left            =   6960
         TabIndex        =   142
         Tag             =   "Click to Select Rehire"
         Top             =   3525
         Visible         =   0   'False
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Rehire         "
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
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataSource      =   " "
         Height          =   285
         Index           =   3
         Left            =   3810
         TabIndex        =   143
         Tag             =   "Termination Cause - Code "
         Top             =   3510
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TECA"
      End
      Begin MSComctlLib.TabStrip tabDates 
         Height          =   855
         Left            =   0
         TabIndex        =   144
         Top             =   0
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   1508
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Employment Dates"
               Key             =   "keyEmpDate"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Pension Dates"
               Key             =   "keyPenDate"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Other Dates"
               Key             =   "keyOtherDate"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   35
         Left            =   4680
         TabIndex        =   193
         Tag             =   "40-Transaction Date"
         Top             =   5070
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1405
         Enabled         =   0   'False
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
         Height          =   195
         Left            =   0
         TabIndex        =   196
         Top             =   5115
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblUserDesc 
         Caption         =   "lblUserDesc"
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
         Left            =   1220
         TabIndex        =   195
         Top             =   5115
         Visible         =   0   'False
         Width           =   1910
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Left            =   3240
         TabIndex        =   194
         Top             =   5115
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgNoSec1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8070
         Picture         =   "festats.frx":096E
         Top             =   5085
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8070
         Picture         =   "festats.frx":0AB8
         Top             =   5085
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Termination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   6870
         TabIndex        =   179
         Top             =   5085
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Caption         =   "Termination Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   0
         TabIndex        =   151
         Top             =   3900
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   150
         Tag             =   "41-Date Terminated"
         Top             =   3225
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   3090
         TabIndex        =   149
         Top             =   3225
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   0
         TabIndex        =   148
         Tag             =   "41-Date Terminated"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblRehired 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Rehired"
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
         Left            =   6750
         TabIndex        =   147
         Tag             =   "41-Date Terminated"
         Top             =   3225
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "              "
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
         Index           =   2
         Left            =   2380
         TabIndex        =   146
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cause"
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
         Left            =   3090
         TabIndex        =   145
         Top             =   3555
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7920
      Top             =   9600
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   9720
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc DataOther 
      Height          =   375
      Left            =   11640
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin MSComDlg.CommonDialog AttachmentDialog 
      Left            =   13680
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEESTATS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim SavEmp, SavOrg, SavPT, SavDOH
Dim XUpdCount
Dim ODate(13), OTermDate
Dim OEmptype  'laura nov 11, 1997 changed from ODate(8)
Dim ODivEdate, ODeptEDate, OPTEDate, OORGEDate
Dim OINTEL, OLANG1, OLANG2, OOMERS, OLDAY
Dim oEmail, OWITHSPOUSE, OEXPYEAR
Dim oFDate, OTDate
Dim oSalDist, oHireCode, OBenGrp
Dim flagFrmLoad As Boolean  'carmen may 00
Dim fglHredsem As String, fglbNew%
Dim savLOA, crtLOA
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim rsDATA2 As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim rsDAT_Other As New ADODB.Recordset
Dim SaveBGroup As String
Dim NewBGroup As String
Dim EmpCountry As String
Dim dtTmpOMERS
Dim strTmpUnion
Dim rsTA As New ADODB.Recordset
Dim oUSER_TEXT1, oUSER_TEXT2, oUSER_NUM1, oUSER_NUM2
Dim locHOOPPBen As Boolean
Dim oPENSIONDATE1, oPENSIONDATE2, oPENSIONDATE3, oPENSIONDATE4, oPENSIONDATE5, oPENSIONDATE6
Dim oOTHERDATE1, oOTHERDATE2, oOTHERDATE3, oOTHERDATE4, oOTHERDATE5
Dim oOTHERDATE6, oOTHERDATE7, oOTHERDATE8, oOTHERDATE9, oOTHERDATE10
Dim OVadim11 As String, OVadim21 As String 'Ticket #19266
Dim xNGSpopFlag As Boolean 'Ticket #19266
Dim oSuperCode 'Ticket #20600 Franks 09/02/2011
Dim fDupEmail_Act
Dim fDupEmail_Term
Dim OVACPC 'Ticket #22710
Dim OSection
Dim locWHRS
Dim MailBody
Dim fglbWDate$, fglbWDateS$

Private Function AUDITSTAT()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsAU2 As New ADODB.Recordset
Dim xDiv, xADD
Dim xBatchID
On Error GoTo AUDIT_ERR

AUDITSTAT = False

rsTB.Open "select ED_DIV,ED_PT FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
  'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
  xDiv = ""
End If
'Number of fields makes * worthwhile ticket# 9899
rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False
 
Dim UpdateAudit As Boolean
UpdateAudit = False
Dim HRChanges As New Collection
Dim UpdateAudit2 As Boolean
UpdateAudit2 = False

'Town of Lasalle - Ticket #23795
'Town of Aurora - Ticket #20931
'Town of Greater Napanee - Ticket #24375
'Ticket #24996 - City of Campbell River
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W" Then
    'Only transfer if Leave of Absent type code
    If IsLOATypeCode(clpCode(1)) Then
        If isChanged_Field(HRChanges, SavEmp, clpCode(1)) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, SavEmp, clpCode(1)) Then UpdateAudit = True
End If

'Ticket #24996 - City of Campbell River - Employee Type not mapped
If glbCompSerial <> "S/N - 2458W" Then
    If isChanged_Field(HRChanges, OEmptype, txtEmpType) Then UpdateAudit = True
End If
If isChanged_Field(HRChanges, SavPT, clpPT) Then UpdateAudit = True

'Town of Aurora - Do not transfer Union Code for Non Union.
If glbCompSerial = "S/N - 2378W" Then
    If clpCode(2).Text <> "0" Then
        If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
    ElseIf SavOrg <> "" And SavOrg <> "0" Then
        If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If SavOrg <> clpCode(2) Then
        If isChanged_Field(HRChanges, "", txtVadim1) Then UpdateAudit = True
    End If
End If
    
'Ticket #23795 - Town of Lasalle - Transfer DOH if First Day is blank and DOH has changed
If glbCompSerial = "S/N - 2379W" And (ODate(8) = dlpDate(0) And ODate(7) <> dlpDate(7) And Not IsDate(dlpDate(0))) Then
    'This causes the ED_FDATE to be passed anyways then in Passing_Changes_Vadim - DOH is passed for blank ED_FDAY
    If isChanged_Field(HRChanges, " ", dlpDate(0)) Then UpdateAudit = True
Else
    If isChanged_Field(HRChanges, ODate(8), dlpDate(0)) Then UpdateAudit = True
End If

'Ticket # 18734 - City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If clpCode(1) = "RNB" Or clpCode(1) = "REB" Or clpCode(1) = "R" Then
        If isChanged_Field(HRChanges, ODate(1), dlpDate(1)) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, ODate(1), dlpDate(1)) Then UpdateAudit = True
End If

'Ticket #25469 - City of Campbell River - No logic behind OMERS date
'Ticket #23795 - Town of Lasalle - No logic behind OMERS date
If glbCompSerial <> "S/N - 2379W" And glbCompSerial <> "S/N - 2458W" Then
    If isChanged_Field(HRChanges, ODate(2), dlpDate(2)) Then UpdateAudit = True
End If

'City of Kawartha Lakes - OMERS Date and RPP# logic
If glbCompSerial = "S/N - 2363W" Then
    If ODate(2) <> dlpDate(2) Then
        If isChanged_Field(HRChanges, "", txtRPP) Then UpdateAudit = True
    End If
End If

'City of Timmins   - RPP# logic
If glbCompSerial = "S/N - 2375W" Then
    If isChanged_Field(HRChanges, "", txtRPP) Then UpdateAudit = True
End If

'Town of Aurora - Ticket #20931 - as per mapping documentation
'City of Niagara Falls - Ticket #20053 - Transfer Benefit Group Code to EMP_CLASS_CODE
If glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2378W" Then
    If isChanged_Field(HRChanges, OBenGrp, clpBGroup) Then UpdateAudit = True
End If

'Ticket #24996 - City of Campbell River - Transfer ED_SECTION to EMP_CLASS_CODE and Benefit Group to EMP_DEFAULT_JOB
If glbCompSerial = "S/N - 2458W" Then
    If isChanged_Field(HRChanges, OBenGrp, clpBGroup) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OSection, clpCode(4)) Then UpdateAudit = True

    'Ticket #28990 - They don't want these to be transferred to Vadim anymore
    'For Sick and Vacation Accruals
    'If isChanged_Field(HRChanges, oUSER_NUM1, txtUserNum1) Then UpdateAudit = True
    'If isChanged_Field(HRChanges, oUSER_NUM2, txtUserNum2) Then UpdateAudit = True
End If

If isChanged_Field(HRChanges, ODate(4), dlpDate(4)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(5), dlpDate(5)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(6), dlpDate(6)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(7), dlpDate(7)) Then UpdateAudit = True

If glbCompSerial <> "S/N - 2375W" Then   'City of Timmins
    If isChanged_Field(HRChanges, ODate(3), dlpDate(3)) Then UpdateAudit = True
End If

If isChanged_Field(HRChanges, ODate(13), dlpDate(12)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(9), dlpDate(8)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(10), dlpDate(9)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(11), dlpDate(10)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(12), dlpDate(11)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODeptEDate, dlpDate(13)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODivEdate, dlpDate(14)) Then UpdateAudit = True
If isChanged_Field(HRChanges, oFDate, dlpDate(15)) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTDate, dlpDate(16)) Then UpdateAudit = True
If isChanged_Field(HRChanges, OINTEL, txtIPHONE) Then UpdateAudit = True
'If isChanged_Field(HRChanges, OLANG1, clpCode(3)) Then UpdateAudit = True
'If isChanged_Field(HRChanges, OLANG2, clpCode(4)) Then UpdateAudit = True
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
    'If isChanged_Field(HRChanges, oHireCode, clpCode(6)) Then UpdateAudit2 = True
    'Move the this function to AUDITSTAT2 on v7.8 'Ticket #15576
Else
    If isChanged_Field(HRChanges, oHireCode, clpCode(6)) Then UpdateAudit = True
End If

'Ticket #23795 - Town of Lasalle - Only transfer Salary Dist if Payment Type is S
If glbCompSerial = "S/N - 2379W" Then
    If getPayType(glbLEE_ID) = "S" Then
        If isChanged_Field(HRChanges, oSalDist, clpSalDist) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, oSalDist, clpSalDist) Then UpdateAudit = True
End If
'If isChanged_Field(HRChanges, oEmail, txtEmail) Then UpdateAudit = True
If isChanged_Field(HRChanges, OWITHSPOUSE, chkSpouse) Then UpdateAudit = True
If isChanged_Field(HRChanges, OEXPYEAR, txtExpYear) Then UpdateAudit = True

Call Passing_Changes(HRChanges, Status, "M", Date, glbLEE_ID)

'This is because Email goes into Client table and above it was trying to update into two different
'tables Employee and Client under same Batch ID as updating the other fields.
Dim HRChanges1 As New Collection
'If not Town of Aurora - Ticket #20931. They do not want to transfer email address as they use if for Pay
'Stub in Vadim. The email address in info:HR will be used for ESS.
If glbCompSerial <> "S/N - 2378W" Then
    If isChanged_Field(HRChanges1, oEmail, txtEmail) Then UpdateAudit = True
End If
Call Passing_Changes(HRChanges1, Status, "M", Date, glbLEE_ID)

'WFC Manulife Interface needs these fields, Ticket #13448
If isChanged_Field(HRChanges, oUSER_TEXT1, txtUserText1) Then UpdateAudit = True
If isChanged_Field(HRChanges, oUSER_TEXT2, txtUserText2) Then UpdateAudit = True

'Ticket #24996 - Not for City of Campbell River because it's being passed above already
If glbCompSerial <> "S/N - 2458W" Then
    If isChanged_Field(HRChanges, oUSER_NUM1, txtUserNum1) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oUSER_NUM2, txtUserNum2) Then UpdateAudit = True
End If

If glbWFC Then 'Ticket #19266
    If isChanged_Field(HRChanges, OVadim11, clpVadim1) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OVadim21, clpVadim2) Then UpdateAudit = True
End If
If glbCompSerial = "S/N - 2382W" Then 'Ticket #20600 Franks 09/02/2011
    If Not (oSuperCode = clpCode(8).Text) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OVadim11, clpVadim1) Then UpdateAudit = True
End If
If glbCompSerial = "S/N - 2417W" Then  'Ticket #22710 - County of Perth
    If isChanged_Field(HRChanges, OVACPC, medVacPPct) Then UpdateAudit = True
End If

If Not UpdateAudit Then GoTo MODNOUPD

MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1": rsTA("AU_LANG2_TABL") = "EDL1"

rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = clpPT.Text
rsTA("AU_DIVUPL") = xDiv

If SavEmp <> clpCode(1).Text Then rsTA("AU_EMP") = clpCode(1).Text
' dkostka - 11/20/2001 - Added 'And txtEmpType <> "-1"' so that it won't try to fill in -1
'   if the field is left blank.
If OEmptype <> txtEmpType And txtEmpType <> "-1" Then rsTA("AU_EMPTYPE") = txtEmpType
If SavPT <> clpPT.Text Then rsTA("AU_PT") = clpPT.Text
If OPTEDate <> dlpDate(34).Text Then If IsDate(dlpDate(34).Text) Then rsTA("AU_PTEDATE") = dlpDate(34).Text
'Ticket #29230 - Daily Vacation Entitlement
If OORGEDate <> dlpDate(36).Text Then If IsDate(dlpDate(36).Text) Then rsTA("AU_ORGEDATE") = dlpDate(36).Text

If OINTEL <> txtIPHONE Then rsTA("AU_INTEL") = txtIPHONE
'If OLANG1 <> clpCode(3).Text Then rsTA("AU_LANG1") = clpCode(3).Text
'If OLANG2 <> clpCode(4).Text Then rsTA("AU_LANG2") = clpCode(4).Text
If glbCompSerial = "S/N - 2296W" Then  'For Essex Library
    If OOMERS <> dlpDate(2).Text And Len(OOMERS) = 0 Then rsTA("AU_PENSION") = "1"
End If
If oEmail <> txtEmail Then rsTA("AU_EMAIL") = txtEmail
If OWITHSPOUSE <> chkSpouse Then rsTA("AU_WITHSPOUSE") = chkSpouse
If OEXPYEAR <> txtExpYear Then
    If Len(txtExpYear) > 0 Then
        rsTA("AU_EXPYEAR") = txtExpYear
    Else
        rsTA("AU_EXPYEAR") = Null
    End If
End If
If SavOrg <> clpCode(2).Text Then
    If Len(clpCode(2).Text) > 0 Then
        rsTA("AU_ORG") = clpCode(2).Text
    Else
        rsTA("AU_ORG") = "-"
    End If
    If glbCompSerial = "S/N - 2217W" Then 'City of Pickering
        'Ticket #20054 Franks 04/04/2011, keep old Union for ADP interface
        If Len(SavOrg) > 0 Then
            rsTA("AU_VADIM2") = Left(SavOrg, 10)
        End If
    End If
End If
If IsDate(dlpDate(0).Text) Then              '12Aug99 js
    If ODate(8) <> dlpDate(0).Text Then      '
        rsTA("AU_FDAY") = dlpDate(0).Text    '
    End If                             '
Else                                   '
    rsTA("AU_FDAY") = Null             '
End If                                 '
If ODate(1) <> dlpDate(1).Text Then
    If Len(dlpDate(1).Text) > 0 Then
        rsTA("AU_LDAY") = dlpDate(1).Text
    Else
        rsTA("AU_LDAY") = CVDate("01/01/01")
    End If
End If
If ODate(2) <> dlpDate(2).Text Then
    If Len(dlpDate(2).Text) > 0 Then
        rsTA("AU_OMDAY") = dlpDate(2).Text
    Else
        rsTA("AU_OMDAY") = CVDate("01/01/01")
    End If
End If

If ODate(3) <> dlpDate(3).Text Then
    If Len(dlpDate(3).Text) > 0 Then
        rsTA("AU_USRDAT1") = dlpDate(3).Text
    Else
        rsTA("AU_USRDAT1") = Null
    End If
End If
If IsDate(dlpDate(4).Text) Then              '12Aug99 js
    If ODate(4) <> dlpDate(4).Text Then      '
        rsTA("AU_UNION") = dlpDate(4).Text   '
    End If                             '
Else                                   '
    rsTA("AU_UNION") = Null            '
End If                                 '
If IsDate(dlpDate(5).Text) Then              '12Aug99 js
    If ODate(5) <> dlpDate(5).Text Then      '
        rsTA("AU_LTHIRE") = dlpDate(5).Text  '
    End If                             '
Else                                   '
    rsTA("AU_LTHIRE") = Null           '
End If                                 '
If IsDate(dlpDate(6).Text) Then              '12Aug99 js
    If ODate(6) <> dlpDate(6).Text Then      '
        rsTA("AU_SENDTE") = dlpDate(6).Text  '
    End If                             '
Else                                   '
    rsTA("AU_SENDTE") = Null           '
End If                                 '
If ODate(7) <> dlpDate(7).Text Then
    rsTA("AU_DOH") = dlpDate(7).Text
End If
If IsDate(dlpDate(8).Text) Then                 '12Aug99 js
    If ODate(9) <> dlpDate(8).Text Then         '
        rsTA("AU_ELIGIBLE") = dlpDate(8).Text   '
    End If                                '
Else                                      '
    rsTA("AU_ELIGIBLE") = Null            '
End If                                    '
If IsDate(dlpDate(9).Text) Then                  '12Aug99 js
    If ODate(10) <> dlpDate(9).Text Then         '
        rsTA("AU_EARLYR") = dlpDate(9).Text      '
    End If                                 '
Else                                       '
    rsTA("AU_EARLYR") = Null               '
End If                                     '
If IsDate(dlpDate(10).Text) Then                  '12Aug99 js
    If ODate(11) <> dlpDate(10).Text Then         '
        rsTA("AU_NORMALR") = dlpDate(10).Text    '
    End If                                  '
Else                                        '
    rsTA("AU_NORMALR") = Null               '
End If                                      '

If IsDate(dlpDate(11).Text) Then                  '12Aug99 js
    If ODate(12) <> dlpDate(11).Text Then         '
        rsTA("AU_LATESTR") = dlpDate(11).Text     '
    End If                                  '
Else                                        '
    rsTA("AU_LATESTR") = Null               '
End If                                      '
If IsDate(dlpDate(12).Text) Then              '11Aug js
    If ODate(13) <> dlpDate(12).Text Then     '
        rsTA("AU_FMLA") = dlpDate(12).Text      '
    End If                              '
Else                                    '
    rsTA("AU_FMLA") = Null              '
End If                                  '
If IsDate(dlpDate(13).Text) Then
    If ODeptEDate <> dlpDate(13).Text Then
        rsTA("AU_DeptEDate") = dlpDate(13).Text
    End If
Else
    rsTA("AU_DeptEDate") = Null
End If
If IsDate(dlpDate(15).Text) Then
    If oFDate <> ODivEdate Then
        rsTA("AU_SFDATE") = dlpDate(15).Text
    End If
Else
    rsTA("AU_SFDATE") = Null
End If

If IsDate(dlpDate(16).Text) Then
    If OTDate <> dlpDate(16).Text Then
        rsTA("AU_STDATE") = dlpDate(16).Text
    End If
Else
    rsTA("AU_STDATE") = Null
End If
If IsDate(dlpDate(14).Text) Then
    If ODivEdate <> dlpDate(14).Text Then
        rsTA("AU_DivEdate") = dlpDate(14).Text
    End If
Else
    rsTA("AU_DivEdate") = Null
End If
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
If glbCompSerial = "S/N - 2296W" Then  'For Essex Library
    If ODate(1) <> dlpDate(1).Text And Len(dlpDate(1).Text) > 0 Then
        rsTA("AU_LDATE") = dlpDate(1).Text
    End If
End If
'Ticket #19067
'If glbCompSerial = "S/N - 2382W" Then  ' Samuel - Ticket #18702
    If IsDate(dlpDate(7).Text) Then
        If CVDate(dlpDate(7).Text) > Date Then
            rsTA("AU_LDATE") = dlpDate(7).Text
        End If
    End If
'End If
If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    If NewHireForms.count > 0 Then 'New Hire only
        If clpCode(1) = "RFT" Or clpCode(1) = "RPT" Then
            rsTA("AU_UIC") = "1"
        Else
            rsTA("AU_UIC") = "2"
        End If
        rsTA("AU_WCBCODE") = "W"
    End If
End If
If oUSER_TEXT1 <> txtUserText1.Text Then
    If Len(txtUserText1.Text) > 0 Then
        rsTA("AU_USER_TEXT1") = txtUserText1.Text
    End If
End If
If oUSER_TEXT2 <> txtUserText2.Text Then
    If Len(txtUserText2.Text) > 0 Then
        rsTA("AU_USER_TEXT2") = txtUserText2.Text
    End If
End If
If oUSER_NUM1 <> txtUserNum1.Text Then
    If Len(txtUserNum1.Text) > 0 And IsNumeric(txtUserNum1.Text) Then
        rsTA("AU_USER_NUM1") = txtUserNum1.Text
    End If
End If
'Ticket #19183
If oUSER_NUM2 <> txtUserNum2.Text Then
    If Len(txtUserNum2.Text) > 0 And IsNumeric(txtUserNum2.Text) Then
        rsTA("AU_USER_NUM2") = txtUserNum2.Text
    End If
End If
If oSalDist <> clpSalDist.Text Then 'Ticket #13828
    If Len(clpSalDist.Text) > 0 Then
        rsTA("AU_SALDIST") = clpSalDist.Text
    End If
End If
If glbWFC Then 'Ticket #19266
        If OVadim11 <> clpVadim1.Text Then rsTA("AU_VADIM1") = clpVadim1.Text
        If OVadim21 <> clpVadim2.Text Then rsTA("AU_VADIM2") = clpVadim2.Text
End If
If glbCompSerial = "S/N - 2382W" Then 'Ticket #20600 Franks 09/02/2011
    If OVadim11 <> clpVadim1.Text Then rsTA("AU_VADIM1") = clpVadim1.Text
End If
If glbCompSerial = "S/N - 2417W" Then  'Ticket #22710
    If OVACPC <> medVacPPct Then
        If IsNumeric(medVacPPct) Then rsTA("AU_VACPC") = medVacPPct * 100
        If IsNumeric(OVACPC) Then rsTA("AU_OLDVAC") = OVACPC * 100
    End If
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = "M"
'If glbSoroc Or glbSyndesis Then
    If Not IsNull(rsDATA("ED_Payroll_ID")) Then rsTA("AU_Payroll_ID") = rsDATA("ED_Payroll_ID")
'End If
rsTA.Update

''If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
''    Call FamilyDayAuditSync(glbLEE_ID, rsTA)
''End If
'Ticket #24729 01/28/2014 Franks - users need to change this form for both ids

If glbWFC Then
    If rsDATA("ED_SECTION") = "TILB" Then
        If Len(glbChgTermReason) > 0 Then
            Call TilburyPayrollIDAudit(rsTA)
        End If
    End If
    If rsDATA("ED_SECTION") = "GREN" Then
        If Len(glbChgTermReason) > 0 Then
            Call WFC_GREN_Audit(rsTA)
        End If
        If Len(glbChgNewEmpnbr) > 0 Then
            Call WFC_GREN_Audit(rsTA)
        End If
    End If
End If
If glbCompSerial = "S/N - 2370W" Then
        If Len(glbChgTermReason) > 0 Then
            Call PayAudit_TermNewhire(rsTA)
        End If
End If
If glbCompSerial = "S/N - 2382W" Then 'Ticket #20600 Franks 09/02/2011
    If Not (oSuperCode = clpCode(8).Text) Then
        Call Samuel_Audit(rsTA, "AU_SUPCODE", clpCode(8).Text, dlpDate(2).Text)
    End If
End If
If rsTA.State <> 0 Then rsTA.Close

Screen.MousePointer = DEFAULT

MODNOUPD:
'If UpdateAudit2 Then
'    rsAU2.Open "SELECT * FROM HRAUDIT2 WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'    rsAU2.AddNew
'    rsAU2("AU_NEWEMP") = "N"
'    rsAU2("AU_TYPE") = "M"
'    rsAU2("AU_COMPNO") = "001"
'    rsAU2("AU_EMPNBR") = glbLEE_ID
'    rsAU2("AU_LDATE") = Date
'    rsAU2("AU_LUSER") = glbUserID
'    rsAU2("AU_LTIME") = Time$
'    rsAU2("AU_UPLOAD") = "N"
'    If oHireCode <> clpCode(6).Text Then
'        rsAU2("AU_HIRECODE") = clpCode(6).Text
'    End If
'    If Len(xDiv) > 0 Then
'        rsAU2("AU_DIVUPL") = xDiv
'    End If
'    rsAU2.Update
'
'End If
AUDITSTAT = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack   '23June99 js
Resume Next
End Function

Private Function chkEStats()
Dim dd As Integer
Dim oCode As String, OCodeD As String
Dim I%, Msg As String, a%
Dim xLOACode1 As Boolean
Dim xLOACode2 As Boolean
Dim xtemDate 'Ticket #20441
'Dim Title$, DgDef, Response%, Msg As String

If glbWFC Then 'Ticket #25785 Franks 07/30/2014 - don't check these if the user is Inquire only
    If gSec_Inq_Basic And Not gSec_Upd_Basic Then
        chkEStats = True
        Exit Function
    End If
End If

chkEStats = False
If glbVadim Then
    'Ticket #25396 - Added If isTransfer(Demographices) because they have Vadim integration but not turned-ON yet
    'calling VadimControl("Check") gives an error
    If isTransfer(Demographices) Then If Not VadimControl("Check") Then Exit Function
End If
'If glbAdv Then  'George Commented for London CCAC changes on Feb 3,2005
'    If Len(dlpDate(3).Text) = 0 Then
'        MsgBox lStr("User Defined is a required field")
'        dlpDate(3).SetFocus
'        Exit Function
'    End If
'End If

If glbWFC Then 'Ticket #21569 Franks 02/22/2012
    If Len(SavEmp) > 0 Then
        If SavEmp <> clpCode(1).Text Then
                'check LOA code
                xLOACode1 = chkLOACode(SavEmp)
                xLOACode2 = chkLOACode(clpCode(1).Text)
                'If Not (xLOACode1 = xLOACode2) Then
                '    If xLOACode1 Or xLOACode2 Then
                'Ticket #22285 Franks 07/16/2012 add "ACP"
                If Not (xLOACode1 = xLOACode2) Or clpCode(1).Text = "ACP" Then
                    If xLOACode1 Or xLOACode2 Or clpCode(1).Text = "ACP" Then
                        Msg = "Cannot change the employment status on this screen." & Chr(10)
                        'If glbWFC Then 'Ticket #21544 Franks 02/07/2012  - for wfc
                            'Msg = Msg & "Please use the Leave & Terminations - Enter a Leave or Return From a Leave. "
                            Msg = Msg & "Please use the appropriate function under Leaves and Terminations to process this type of transaction."
                        'End If
                        MsgBox Msg
                        clpCode(1).SetFocus
                        Exit Function
                    End If
                End If
        End If
    End If
End If


If glbLinamar Then
    If SavEmp <> clpCode(1).Text Then
        If SavEmp = "TEMP" Or clpCode(1).Text = "TEMP" Then
            MsgBox "Please go to Termporary Lay-Off procedure to Change the Employment Status"
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
    If Not chkElig Then
        If Len(dlpDate(8)) = 0 And gSec_Show_DOB Then
            tabDates.SelectedItem = tabDates.Tabs(2)
            fraDatePension.Visible = gSec_Show_DOB
            MsgBox lStr("Eligibility date is requried field")
            dlpDate(8).SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(6)) = 0 Then
        MsgBox lStr("Hire Code is requried field")
        clpCode(6).SetFocus
        Exit Function
    End If
Else
    If savLOA <> 0 Or chkLeave <> 0 Then
        MsgBox "Please go to Employee Leaves of Absence procedure to Change the Employment Status"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(1).Text) < 1 Then  'Or lblCodeDesc(1) = "Unassigned" Then
    MsgBox "Employment Status must be entered"
    clpCode(1).SetFocus '17Aug99 js
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Employment Status code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

'Ticket #22682 - Release 8.0 - For everyone if Employment Status is marked as LOA - do not allow to save. Use
'Enter LOA' under Leaves and Termination menu to process leave.
'Ticket #18235 - Cannot change to LOA type Status - Samuel, Son & Co., Limited
'If glbCompSerial = "S/N - 2382W" Then
    If SavEmp <> clpCode(1).Text Then
        xLOACode1 = chkLOACode(clpCode(1).Text)
        If xLOACode1 Then
            Msg = "Employment Status cannot be changed to Leave of Absence type on this screen." & Chr(10)
            Msg = Msg & "Please use the Leave and Terminations: 'Enter a Leave' or 'Return From a Leave' screens. "
            MsgBox Msg, vbInformation, "Employment Status - Leave of Absence Type"
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
'End If

If Not glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    If Len(clpPT.Text) < 1 Then
        MsgBox lblPT.Caption & " must be entered"
        clpPT.SetFocus
        Exit Function
    End If
End If
If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If

If Len(dlpDate(34).Text) > 0 Then
    If Not IsDate(dlpDate(34).Text) Then
        MsgBox "Invalid " & lStr("Category") & " Effective Date"
        dlpDate(34).SetFocus
        Exit Function
    Else
        If Len(clpPT.Text) = 0 Or clpPT.Caption = "Unassigned" Then
            MsgBox lStr("Category code must be valid")
            clpPT.SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    If Len(clpCode(7).Text) = 0 Then
        MsgBox lStr("Region") & " must be entered"
        clpCode(7).SetFocus
        Exit Function
    Else
        If clpCode(7).Caption = "Unassigned" Then
            MsgBox lStr("Region") & " must be valid"
            clpCode(7).SetFocus
            Exit Function
        End If
    End If
End If

'wellington duffrine ticket ##17736
'2409W Delisle Youth Services Ticket #27798 Franks 09/26/2016
If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2409W" Then
    If Len(clpCode(2).Text) <= 0 Then
        MsgBox lblUnion.Caption & " must be entered"
        clpCode(2).SetFocus
        Exit Function
        
    End If
    If Len(clpCode(2).Text) > 0 Then
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox lStr("Union code must be valid")
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
Else
    If Len(clpCode(2).Text) > 0 Then
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox lStr("Union code must be valid")
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
    'Ticket #29230 - Union Effective Date validation
    If Len(dlpDate(36).Text) > 0 Then
        If Not IsDate(dlpDate(36).Text) Then
            MsgBox "Invalid " & lStr("Union") & " Effective Date"
            dlpDate(36).SetFocus
            Exit Function
        Else
            If Len(clpCode(2).Text) = 0 Or clpCode(2).Caption = "Unassigned" Then
                MsgBox lStr("Union code must be valid")
                clpCode(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If

''Ticket #25040 - Remove the hiding of the Salary Distribution field.
'Ticket #24543 - Macaulay Child Development Centre
'If glbCompSerial <> "S/N - 2420W" Then
    If clpSalDist.Caption = "Unassigned" Then
        MsgBox lblSalDist.Caption & " must be valid"
        clpSalDist.SetFocus
        Exit Function
    End If
'End If

'Simona - begin -Assessement Strategies - # 14963
If glbCompSerial = "S/N - 2401W" Then
    If Len(txtIPHONE.Text) < 1 Then
        MsgBox lblIPhone.Caption & " must be entered"
            txtIPHONE.SetFocus
            Exit Function
        End If
    
        If Len(txtEmail.Text) < 1 Then
        MsgBox lblEmail.Caption & " must be entered"
            txtEmail.SetFocus
            Exit Function
        End If
    
End If

'Granite Club
If glbCompSerial = "S/N - 2241W" Then
    If Len(comEmpType) < 1 Then
        MsgBox "Employment Type must be selected"
        comEmpType.SetFocus
        Exit Function
    End If

    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If (glbCompSerial = "S/N - 2347W") Then
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If (glbCompSerial = "S/N - 2259W") Or (glbCompSerial = "S/N - 2394W") Then 'County of Oxford 'St. John's Ticket #14752
    If Len(dlpDate(5).Text) < 1 Then
        MsgBox lStr("Last Hire is a required field")
        dlpDate(5).SetFocus
        Exit Function
    End If
End If
If (glbCompSerial = "S/N - 2394W") Then 'St. John's Ticket #14752
    If OEmptype <> txtEmpType And txtEmpType <> "-1" Then
        If ODate(5) = dlpDate(5).Text Then
            Msg = "The MediPay Status has been changed from '" & OEmptype & "' to '" & txtEmpType & "', " & Chr(10)
            Msg = Msg & "but the " & lStr("Last Hire") & " has not been changed. " & Chr(10)
            Msg = Msg & "Please enter a new " & lStr("Last Hire") & "."
            MsgBox Msg
            dlpDate(5).SetFocus
            Exit Function
        End If
    End If
End If
If glbCompSerial = "S/N - 2375W" Then  'For Timmis
    If Len(clpCode(2)) = 0 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(dlpDate(6)) = 0 Then
        MsgBox lStr("Seniority Date is a required field")
        dlpDate(6).SetFocus
        Exit Function
    End If
    If Len(dlpDate(0)) = 0 Then
        MsgBox lStr("First Day is a required field")
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Len(dlpDate(2)) = 0 And clpPT = "FT" Then
        MsgBox lStr("OMERS Date is a required field")
        dlpDate(2).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2357W" And Data1.Recordset("ED_COUNTRY") = "CANADA" Then   'I.T. Xchange
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2394W" Then ' St. John's Rehab Hospital - Ticket #14572
    If Len(Trim(clpBGroup.Text)) = 0 Then
        MsgBox lStr("Benefit Group is a required field")
        clpBGroup.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2453W" Then  'Town of Gander Ticket #24518 Franks 06/04/2015
    If Len(Trim(clpCode(2).Text)) = 0 Then
        MsgBox lStr("Union") & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(Trim(clpBGroup.Text)) = 0 Then
        MsgBox lStr("Benefit Group is a required field")
        clpBGroup.SetFocus
        Exit Function
    End If
End If

'Hemu
If Len(clpBGroup.Text) > 0 Then
    If clpBGroup.Caption = "Unassigned" Then
        MsgBox lStr("Benefit Group must be valid")
        clpBGroup.SetFocus
        Exit Function
    End If
End If
'Hemu

If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2466W" Then    ' St. John's Rehab Hospital - Ticket #14572
    If Len(clpSalDist.Text) < 1 Then
        MsgBox lStr("Salary Distribution is a required field")
        clpSalDist.SetFocus
        Exit Function
    End If
End If
'Ticket #18844 Franks 01/13/2011  Town of Orangeville
'Ticket #20666 Franks 07/19/2011  '2429W Municipality of North Perth
'If glbCompSerial = "S/N - 2383W" Or glbCompSerial = "S/N - 2429W" Then
'Ticket #23189 Franks 02/07/2013 - removed this for Orangeville
'2436W Family Day Ticket #24729 01/20/2014
'Ticket #25376 - Community Living Access Support Services
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2436W" Or glbCompSerial = "S/N - 2301W" Then
    If Len(clpSalDist.Text) < 1 Then
        MsgBox lStr("Salary Distribution") & " is a required field"
        clpSalDist.SetFocus
        Exit Function
    End If
End If

'Hemu - Begin - Check for validity of Location Code since it has been moved from
'               Demographics screen. Checking for validity of the code in Demographics
'               screen has been disabled. Ticket # 4972
If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    If Len(clpCode(0).Text) > 0 Then
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox lblTitle(20).Caption & " must be valid"
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
End If
'Hemu - End

If glbCompSerial = "S/N - 2410W" Then  'Ticket #18603 - Frontenac
    If Len(clpCode(2).Text) = 0 Then 'Ticket #19891
            MsgBox lblUnion.Caption & " is a required field"
            clpCode(2).SetFocus
            Exit Function
    End If
    If Len(clpCode(0).Text) > 0 Then
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox lblTitle(20).Caption & " must be valid"
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(clpBGroup.Text)) = 0 Then
        MsgBox lStr("Benefit Group is a required field")
        clpBGroup.SetFocus
        Exit Function
    End If
    'Ticket #19071
    If Len(Trim(clpSalDist.Text)) = 0 Then
        MsgBox lblSalDist.Caption & " is a required field"
        clpSalDist.SetFocus
        Exit Function
    End If
    If Len(Trim(clpCode(0).Text)) = 0 Then
        MsgBox lblTitle(20).Caption & " is a required field"
        clpCode(0).SetFocus
        Exit Function
    End If
End If

'For i% = 3 To 4
'    If Len(clpCode(i).Text) > 0 Then
'        If clpCode(i).Caption = "Unassigned" Then
'            MsgBox "Language Code must be valid"
'            clpCode(i).SetFocus
'            Exit Function
'        End If
'    End If
'Next i%
' MC - dkostka - 05/07/2001 - Below code was supposted to be ONLY for Linamar.  Was put in for everyone by mistake.
If Len(comEmpType) < 1 And glbLinamar Then
    MsgBox "Employment Type must be selected"
    comEmpType.SetFocus
    Exit Function
End If
'' end dkostka - 05/07/2001
'If Len(clpPT.Text) < 2 Then
'    MsgBox "Category must be entered"
'     clpPT.SetFocus
'    Exit Function
'End If

If glbCompSerial = "S/N - 2394W" Then ' St. John's Rehab Hospital - Ticket #14572
    If Len(dlpDate(17).Text) < 1 Then
        MsgBox lStr(lblUserText2.Caption) & " is a required field"
        dlpDate(17).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDate(17).Text) Then
        MsgBox lStr(lblUserText2.Caption) & " is not a valid date"
        dlpDate(17).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(7).Text) < 1 Then
    tabDates.SelectedItem = tabDates.Tabs(1)
    fraDateEmp.Visible = True
    MsgBox lStr("Original Hire Date must be entered")
    dlpDate(7).SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate(7).Text) Then
        MsgBox lStr("Original Hire Date is not a valid date")
        dlpDate(7).SetFocus
        Exit Function
    Else
        If IsDate(lblDob) Then
            If DaysBetween(lblDob, dlpDate(7).Text) < 1 Then
                MsgBox "Date can not be prior to individuals Birth Date"
                dlpDate(7).SetFocus
                Exit Function
            End If
        Else
            MsgBox "This employee's Birth Date is missing on the Demographics screen. " & lStr("Original Hire Date") & " cannot be validated."
            Exit Function
        End If
    End If
End If
If glbLinamar Then
    If Len(dlpDate(6).Text) < 1 Then
        tabDates.SelectedItem = tabDates.Tabs(1)
        fraDateEmp.Visible = True
        MsgBox lStr("Seniority Date must be entered")
        dlpDate(6).SetFocus
        Exit Function
    End If
    If NewHireForms.count > 0 Then
        If Len(dlpDate(13).Text) < 1 Then
            tabDates.SelectedItem = tabDates.Tabs(1)
            fraDateEmp.Visible = True
            MsgBox lStr("Department Start Date must be entered")
            dlpDate(13).SetFocus
            Exit Function
        End If
    End If
End If
      
    
'.....and Granite Club
If glbCompSerial = "S/N - 2297W" Or glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2241W" Then
    If Len(dlpDate(6).Text) < 1 Then
        MsgBox lStr("Seniority Date must be entered")
        dlpDate(6).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2366W" Then
    If clpCode(1) = "PERM" Or clpCode(1) = "MAT" Then
        If Not (clpPT = "FT" Or clpPT = "PT") Then
            MsgBox lStr("Category must be 'FT' or 'PT' ") & lStr("if Employment Status is 'PERM' or 'MAT' ")
            clpPT.SetFocus
            Exit Function
        End If
    End If
End If

If glbLambton Then
    If Len(dlpDate(0).Text) < 1 Then
        MsgBox lStr("First Day must be entered")
        dlpDate(0).SetFocus
        Exit Function
    End If
End If

For I% = 6 To 4 Step -1
    If Len(dlpDate(I%).Text) > 0 Then
        If Not IsDate(dlpDate(I%).Text) Then
            If I% = 6 Then
                MsgBox lStr("Seniority date entered is not a valid date!")
            ElseIf I% = 5 Then
                MsgBox lStr("Last Hire date entered is not a valid date!")
            ElseIf I% = 4 Then
                MsgBox "Union Date entered is not a valid date!"
            End If
            dlpDate(I%).SetFocus
            Exit Function
        Else
            If IsDate(lblDob) Then
                If DaysBetween(lblDob, dlpDate(I%).Text) < 1 Then
                    MsgBox "Date can not be prior to individuals Birth Date"
                    dlpDate(I%).SetFocus
                    Exit Function
                End If
            Else
                MsgBox "This employee's Birth Date is missing on the Demographics screen. Dates cannot be validated."
                Exit Function
            End If
        End If
    End If
Next I%
'commented by Bryan Ticket#11594
'Commented in 7.4 Jerry doesn't know why this was uncommmented in 7.6
'If Val(txtExpYear.Text) <> 0 Then
'    If Year(Date) - Val(txtExpYear) > 100 Or Val(txtExpYear) > Year(Date) Then
'        MsgBox "Experience Year is not a valid year"
'        txtExpYear.SetFocus
'        Exit Function
'    End If
'End If

For I% = 0 To 3
    If Len(dlpDate(I%).Text) > 0 Then
        If Not IsDate(dlpDate(I%).Text) Then
            If I% = 0 Then
                MsgBox "First Day entered is not a valid date!"
            ElseIf I% = 1 Then
                MsgBox "Last Day entered is not a valid date!"
            ElseIf I% = 2 Then
                MsgBox "OMERS Date entered is not a valid date!"
            ElseIf I% = 3 Then
                MsgBox lStr("User Defined date") & " entered is not a valid date!"
            End If
            dlpDate(I%).SetFocus
            Exit Function
        Else
            If IsDate(lblDob) Then
                If DaysBetween(lblDob, dlpDate(I%).Text) < 1 Then
                    MsgBox "Date can not be prior to individuals Birth Date"
                    dlpDate(I%).SetFocus
                    Exit Function
                End If
            Else
                MsgBox "This employee's Birth Date is missing on the Demographics screen. Dates cannot be validated."
                Exit Function
            End If
        End If
    End If
Next I%

'Ticket #24203 - Family Day Care Services
'Ticket #21504 - Kerry's Place
If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2436W" Then
    If Len(clpBGroup.Text) > 0 Then
        If SaveBGroup <> clpBGroup.Text Then
            If Not IsDate(dlpDate(3).Text) Then
                MsgBox lStr("User Defined date") & " is required when Benefit Group entered for Benefit Effective Date."
                dlpDate(3).SetFocus
                Exit Function
            End If
        End If
    End If
End If

For I% = 8 To 12    '23June99 js - changed 8 to "11" to 8 to "12"
    If Len(dlpDate(I%).Text) > 0 Then
        If Not IsDate(dlpDate(I%).Text) Then
            If I% = 8 And gSec_Show_DOB Then
                MsgBox "Eligibility date entered is not a valid date!"
            ElseIf I% = 9 And gSec_Show_DOB Then
                MsgBox "Earliest Retirement date entered is not a valid date!"
            ElseIf I% = 10 And gSec_Show_DOB Then
                MsgBox "Normal Retirement date entered is not a valid date!"
            ElseIf I% = 11 And gSec_Show_DOB Then
                MsgBox "Latest Retirement date entered is not a valid date!"
            ElseIf I% = 12 Then     '23June99 - added condition and message
                MsgBox "FMLA date entered is not a valid date!"
            End If
            dlpDate(I%).SetFocus
            Exit Function
        End If
    End If
Next I%
If Not glbLinamar Then
    If Len(dlpDate(14).Text) > 1 Then
        If Not IsDate(dlpDate(14).Text) Then
            MsgBox lStr("Invalid Effictive Date for Division")
            dlpDate(14).SetFocus
            Exit Function
        End If
    End If
End If
If Len(dlpDate(13).Text) > 1 Then
    If Not IsDate(dlpDate(13).Text) Then
        MsgBox "Invalid Effictive Date for Department"
        dlpDate(13).SetFocus
        Exit Function
    End If
End If
If glbtermopen Then
    If Not IsDate(dlpTermDate.Text) Then
        MsgBox "Invalid Termination Date"
        dlpTermDate.SetFocus
        Exit Function
    End If
    If Len(clpCode(5).Text) = 0 Then
        MsgBox "Invalid Termination Reason"
        clpCode(5).SetFocus
        Exit Function
    End If
    If clpCode(5).Caption = "Unassigned" Then
        MsgBox "Invalid Termination Reason"
        clpCode(5).SetFocus
        Exit Function
    End If
    If Len(dlpRehired.Text) > 0 Then
      If Not IsDate(dlpRehired.Text) Then
        MsgBox "Invalid Date Rehired"
        dlpRehired.SetFocus
        Exit Function
      End If
    End If
End If
'For Essex Library May 23,2002 Franks
If Not glbtermopen Then
    If glbCompSerial = "S/N - 2296W" Then
        If ODate(1) <> dlpDate(1).Text And Len(dlpDate(1).Text) > 0 Then
            Msg = "Are you sure you want to terminate this employee? "
            a% = MsgBox(Msg, 36, "Confirm ")
            If a% <> 6 Then
                dlpDate(1).SetFocus
                Exit Function
            End If
        End If
    End If
End If
'For Essex Library May 23,2002 Franks

If glbWFC Then 'Ticket #21626 Franks 02/27/2012
    If Not IsDate(dlpDate(15).Text) Then
            MsgBox ("From Date cannot be blank. If the From Date is not known, use the Original Date of Hire.")
            dlpDate(15).SetFocus
            Exit Function
    End If
    'If glbEmpCountry = "U.S.A." Then 'Ticket #23564 Franks 04/15/2013
        If Not IsDate(dlpDate(34).Text) Then
                MsgBox lblPTEDate.Caption & (" cannot be blank. If the " & lblPTEDate.Caption & " is not known, use the Original Date of Hire.")
                dlpDate(34).SetFocus
                Exit Function
        End If
        If Not (SavPT = clpPT.Text) Then 'Ticket #23564 Franks 04/15/2013
            '"   If the Effective Date is populated and they make a change to the Category
            'but do not enter a new Effective Date, make the Effective Date equal to "Today".
            If NewHireForms.count = 0 Then 'Ticket #23837 Franks 05/28/2013 - modify only
                If OPTEDate = dlpDate(34).Text Then
                    If IsDate(OPTEDate) Then
                        dlpDate(34).Text = Date
                    End If
                End If
            End If
        End If
    'End If
End If
If Len(dlpDate(15).Text) > 0 Then
    If Not IsDate(dlpDate(15).Text) Then
        MsgBox "Invalid From Date for Employment Status"
        dlpDate(15).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(16).Text) > 0 Then
    If Not IsDate(dlpDate(16).Text) Then
        MsgBox "Invalid To Date for Employment Status"
        dlpDate(16).SetFocus
        Exit Function
    End If
End If

If IsDate(dlpDate(15).Text) And IsDate(dlpDate(16).Text) Then       'Serbo
    If DaysBetween(dlpDate(15), dlpDate(16).Text) < 1 Then
        MsgBox "Status To date can not be prior to Status From date"
        dlpDate(16).SetFocus
        Exit Function
    End If
End If

If Len(txtExpYear.Text) <> 0 Then
    If Not IsNumeric(txtExpYear.Text) Then
        MsgBox "Experience Year must be numeric"
        txtExpYear.SetFocus
        Exit Function
    End If
End If

'If Len(dlpDate(7).Text) > 0 And Len(dlpDate(0).Text) > 0 Then
'    If DaysBetween(dlpDate(7).Text, dlpDate(0).Text) < 0 Then
'        MsgBox "First Day can not be prior to Original Hire date"
'        dlpDate(0).SetFocus
'        Exit Function
'    End If
'End If

If glbSamuel Then 'Ticket #23464 Franks 03/22/2013
    'they don't want this logic
Else
    If Len(dlpDate(1).Text) > 0 And Len(dlpDate(0).Text) > 0 Then
        If DaysBetween(dlpDate(0).Text, dlpDate(1).Text) < 0 Then
            MsgBox "Last Day can not be prior to First Day"
            dlpDate(1).SetFocus
            Exit Function
        End If
    End If
End If

If glbGP And Left(clpCode(1), 1) = "I" Then
    If Len(dlpDate(15)) = 0 Then
        MsgBox lStr("From Date must be a valid date"), vbInformation + vbOKOnly
        dlpDate(15).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Ticket #14752
    If clpCode(1).Text = "TERM" Then
        If Len(dlpDate(1).Text) = 0 Then ' Last Day
            MsgBox lStr("Last Day") & " is required if Employment Status is 'TERM'", vbInformation + vbOKOnly
            dlpDate(1).SetFocus
            Exit Function
        End If
        txtEmpType.Text = "X"
    End If
End If

'Ticket #21376 - Charton Hobbs
If glbCompSerial = "S/N - 2418W" Then
    If Len(Trim(txtEmail.Text)) = 0 Then
        MsgBox "Email Address cannot be blank"
        txtEmail.SetFocus
        Exit Function
    End If
    If Len(Trim(txtUserText1.Text)) = 0 Then
        MsgBox lblUserText1.Caption & " cannot be blank"
        txtUserText1.SetFocus
        Exit Function
    End If
    If Len(dlpDate(3).Text) = 0 Or IsDate(dlpDate(3).Text) = False Then
        MsgBox lStr("User Defined date cannot be blank")
        dlpDate(3).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2214W" Then
    If Len(txtEmail.Text) = 0 Then
        MsgBox "Email address must be filled in"
        txtEmail.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        clpCode(2).Text = "ATT1"
        Exit Function
    End If
    If Len(clpCode(6).Text) = 0 Then
            MsgBox lStr("Hire Code is requried field")
            clpCode(6).SetFocus
            Exit Function
    End If
    If Not IsDate(dlpDate(15)) Then
            MsgBox ("From Date is requried field")
            dlpDate(15).SetFocus
            Exit Function
    End If
    If clpPT.Text = "FT" Then
        If Len(txtIPHONE.Text) = 0 Then
                MsgBox lStr("Internal Phone Extension is required field")
                txtIPHONE.SetFocus
                Exit Function
        End If
    End If
    If locHOOPPBen Then
        If Not IsDate(dlpDate(2)) Then
                MsgBox lStr("OMERS Date is requried field")
                dlpDate(2).SetFocus
                Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2174W" Then 'Kawartha-Haliburton CAS 'Ticket #23382 Franks 04/09/2013
    If Not IsDate(dlpDate(15)) Then
            MsgBox ("From Date is requried field")
            dlpDate(15).SetFocus
            Exit Function
    End If
End If

If glbWFC Then
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Union code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDate(36).Text) Then
        MsgBox lStr("Union") & " Effective Date' is required"
        Exit Function
    End If
    
    ''Ticket #23903 Franks 06/19/2013 - moved this section to the bottom
    '''Ticket #22584 Franks 10/03/2012 "   Do not allow the user to change the UNION
    ''If Len(SavOrg) > 0 And NewHireForms.count = 0 Then
    ''    If Not (SavOrg = clpCode(2).Text) Then
    ''        MsgBox lStr("Union") & " code cannot be changed on this screen." & Chr(10) & "Use the Transfer Out/IN screens to change an employee's Union Code."
    ''        clpCode(2).SetFocus
    ''        Exit Function
    ''    End If
    ''End If
    
    If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
        If Len(txtEmail.Text) = 0 Then
            MsgBox "Email address must be filled in"
            Call WFCDefaultEmailDisp 'Ticket #25275 Franks 04/02/2014
            txtEmail.SetFocus
            Exit Function
        End If
        If Len(txtIPHONE.Text) = 0 Then
            MsgBox "Internal Telephone must be filled in"
            txtIPHONE.SetFocus
            Exit Function
        End If
    End If
    If Len(txtUserText1.Text) > 0 Then
        'Ticket #22776 Franks 11/06/2012 - begin
        'If Len(txtUserText2.Text) = 0 Then
        '    MsgBox "If " & lStr(lblUserText1.Caption) & " was entered then " & lStr(lblUserText2.Caption) & " is required."
        '    txtUserText2.SetFocus
        '    Exit Function
        'End If
        If Len(comUserText2.Text) = 0 Then
            MsgBox "If " & lStr(lblUserText1.Caption) & " was entered then " & lStr(lblUserText2.Caption) & " is required."
            comUserText2.SetFocus
            Exit Function
        End If
        'Ticket #22776 Franks 11/06/2012 - end
        If Len(txtUserNum1.Text) = 0 Then
            MsgBox "If " & lStr(lblUserText1.Caption) & " was entered then " & lStr(lblUserNum1.Caption) & " is required."
            txtUserNum1.SetFocus
            Exit Function
        End If
    End If
    'Ticket #19266 Franks 12/13/2010
    'move this validation check to from Banking screen
    If glbEmpCountry = "U.S.A." Then
        If Len(clpVadim2.Text) < 1 Then
            MsgBox lStr("Vadim Field 2 is required field")
            clpVadim2.SetFocus
            Exit Function
        ElseIf Len(clpVadim2.Text) > 0 And clpVadim2.Caption = "Unassigned" Then
            MsgBox lStr("Vadim Field 2 must be valid")
            clpVadim2.SetFocus
            Exit Function
        End If
    Else
        'Ticket #20049 Franks 05/30/2011
        If Not IsNull(rsDATA("ED_SECTION")) Then
            If rsDATA("ED_SECTION") = "MISS" Or rsDATA("ED_SECTION") = "KIPL" Then
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
        End If
    End If
    'WFC-If Employment Status = PROB, eligible for Pension must be NO
    'Pension Change - August 10-2010.docx
    If clpCode(1).Text = "PROB" Then
        If comEmpType.Text = "Y - Yes" Then
            MsgBox lblEEType.Caption & " must be No if Employment Status is 'PROB."
            comEmpType.SetFocus
            Exit Function
        End If
    Else
        'If Employment Status was changed from PROB to other, and County of Employment = CANADA
        'and Eligible for Pension = N, display a message
        If SavEmp = "PROB" Then
            If comEmpType.Text = "N - No" And glbEmpCountry = "CANADA" Then
                Msg = "Is this employee eligible for pension?"
                a% = MsgBox(Msg, 36, "Confirm ")
                If a% = 6 Then
                    comEmpType.ListIndex = 0
                    DoEvents
                End If
            End If
        End If
    End If
    'Ticket #24936 Franks 02/05/2014 - "   If union = "U838" and DOH >= July 1, 2011, Eligible for Pension cannot be Y.
    If clpCode(2).Text = "U838" Then
        If comEmpType.Text = "Y - Yes" Then
            If IsDate(dlpDate(7).Text) Then
                If CVDate((dlpDate(7).Text)) >= CVDate("Jul 1, 2011") Then
                    MsgBox lblEEType.Caption & " must be No if Union is 'U838' and Original Hire Date >= 'July 1, 2011'."
                    comEmpType.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If NewHireForms.count > 0 Then
        If Not IsDate(dlpDate(15)) Then
            'MsgBox ("From Date is requried field")
            'dlpDate(15).SetFocus
            'Exit Function
            dlpDate(15).Text = dlpDate(7).Text
        End If
        'Ticket #19266 Frank 10/27/2010, new hire
        'For Salaried employees, the "Other Date 1" defaults to the Original Date of Hire. Salaried is defined if union code is equal to "NONE" or "EXEC"
        If glbEmpCountry = "U.S.A." Then 'Ticket #20387 Franks 05/27/2011
            Call WFCNGSStartDate 'Ticket #24695 Franks 11/28/2013
        End If
    Else
        If SavEmp <> clpCode(1).Text Then
            If oFDate = dlpDate(15).Text Then
                Msg = "The Employment Status has been changed from '" & SavEmp & "' to '" & clpCode(1).Text & "', " & Chr(10)
                Msg = Msg & "but the " & ("From Date") & " has not been changed. " & Chr(10)
                Msg = Msg & "Please enter a new " & lStr("From Date") & "."
                MsgBox Msg
                clpCode(1).SetFocus
                Exit Function
            End If
        End If
        'Ticket #18804 - remove this logic
        'If Not (ODate(7) = dlpDate(7).Text) Then
        '    glbAccessPswd = False
        '    frmAccessPswd.Show 1
        '    If glbAccessPswd = False Then   'Access Denied
        '        MsgBox "Can not change Original Hire Date."
        '        dlpDate(7).SetFocus
        '        Exit Function
        '    End If
        'End If
    End If
    
    'Ticket #16395 09/11/2009 Frank
    If comEmpType.Text = "Y - Yes" Then
        If NewHireForms.count > 0 Then 'New Hire only
            If Len(dlpDate(8)) = 0 And gSec_Show_DOB Then
                'On New Hire Status/Dates screen, if Pension Eligibility = Yes,
                'default the Membership Entry Date to equal the DOH
                dlpDate(8).Text = dlpDate(7).Text
            End If
        Else
            If Len(dlpDate(8)) = 0 And gSec_Show_DOB Then
                MsgBox lStr("Eligibility date is requried field")
                'dlpDate(8).SetFocus
                Exit Function
            End If
        End If
        If IsDate(oFDate) And IsDate(oFDate) Then
        'If the FROM DATE is changed and the Pension Eligibility is Y, display a message saying
            If Len(dlpDate(15).Text) > 0 Then 'Ticket #21569 Franks 02/22/2012
                If Not (CVDate(oFDate) = CVDate(dlpDate(15))) Then
                    Msg = "A change in the From Date may affect the employee's Credited Service or Earned Pension." & Chr(10)
                    Msg = Msg & "Please go into the Pension Master and make the necessary changes to " & Chr(10)
                    Msg = Msg & "the Effective Status Date, Earned Pension and/or Credited Service."
                    MsgBox Msg
                End If
            End If
        End If
        'Ticket #19678 Franks 01/24/2011 - for the Special Early Retirement function
        If comEmpType.Text = "Y - Yes" Then
            If Len(clpCode(2).Text) > 0 Then 'Union
                If Not IsNull(rsDATA("ED_SECTION")) Then
                    If Len(clpCode(6).Text) = 0 Then
                        If (rsDATA("ED_SECTION") = "TILB" And clpCode(2).Text = "C127") Or (rsDATA("ED_SECTION") = "WHBY" And clpCode(2).Text = "C222") Then
                            clpCode(6).Text = "Y"
                        Else
                            clpCode(6).Text = "N"
                        End If
                    End If
                End If
            End If
        End If
    End If
    'Ticket #23361 Franks 03/11/2013 - begin
    '"   Cannot change Eligible for Pension to "N" if the current status  is 'C', 'D', 'N' or 'T' .
    If NewHireForms.count = 0 Then 'change only
        If comEmpType.Text = "N - No" Then
            If Left(OEmptype, 1) = "Y" Then 'from Y to N
                Dim xLocPenStatu As String
                xLocPenStatu = getWFCPenStatus(glbLEE_ID)
                If xLocPenStatu = "C" Or xLocPenStatu = "D" Or xLocPenStatu = "N" Or xLocPenStatu = "T" Then
                    MsgBox "Cannot change " & lblEEType.Caption & " to 'N' if Pension Status is 'C', 'D', 'N' or 'T'."
                    comEmpType.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    'Ticket #23361 Franks 03/11/2013 - end
    If Len(SavEmp) > 0 Then
        If SavEmp <> clpCode(1).Text Then
            If Not glbWFC Then
                If SavEmp = "LOA" Or clpCode(1).Text = "LOA" Then
                    If SavEmp = "LOA" Then
                        MsgBox "To return an Employee back to an active status, " & Chr(10) & "please use 'Return from a Leave' under Leaves and Terminations."
                    End If
                    If clpCode(1).Text = "LOA" Then
                        MsgBox "To enter an Employee going on a Leave, " & Chr(10) & "please use 'Enter a Leave' under Leaves and Terminations"
                    End If
                    clpCode(1).SetFocus
                    Exit Function
                End If
            Else  'If glbWFC Then
                ''Ticket #21569 Franks 02/22/2012, move this function to the top
                '''check LOA code
                ''xLOACode1 = chkLOACode(SavEmp)
                ''xLOACode2 = chkLOACode(clpCode(1).Text)
                ''If Not (xLOACode1 = xLOACode2) Then
                ''    If xLOACode1 Or xLOACode2 Then
                ''        Msg = "Cannot change the employment status on this screen." & Chr(10)
                ''        'If glbWFC Then 'Ticket #21544 Franks 02/07/2012  - for wfc
                ''            'Msg = Msg & "Please use the Leave & Terminations - Enter a Leave or Return From a Leave. "
                ''            Msg = Msg & "Please use the appropriate function under Leaves and Terminations to process this type of transaction."
                ''        'End If
                ''        MsgBox Msg
                ''        clpCode(1).SetFocus
                ''        Exit Function
                ''    End If
                ''End If
            End If
        End If
    End If
    
    If glbEmpCountry = "U.S.A." Then
    'Ticket #22663 Franks 10/16/2012
    'display a message saying "Employee must be FT in order to have a NGS Sub-Group.".
        If Len(clpVadim1.Text) > 0 Then
            'If Not clpPT.Text = "FT" Then
            If Not (clpPT.Text = "FT" Or clpPT.Text = "PT") Then 'Ticket #22991 Franks 12/24/2012
                MsgBox "Employee must be FT or PT in order to have a NGS Sub-Group"
                Exit Function
            End If
        End If
    End If
    
    'Ticket #21544 Franks 02/07/2012
    If NewHireForms.count = 0 Then 'not New Hire
        If Not OVadim11 = clpVadim1.Text Then
            If Len(dlpDate(25).Text) = 0 Then 'Ticket #21569 Franks 02/22/2012
                Msg = lblVadim11.Caption & " was changed from " & OVadim11 & " to " & IIf(Len(clpVadim1.Text) = 0, "blank", clpVadim1.Text) & " " & Chr(10)
                Msg = Msg & "Do you want to enter the NGS End Date? "
                a% = MsgBox(Msg, 36, "Confirm ")
                If a% = 6 Then
                    tabDates.Tabs(3).selected = True
                    Call tabDates_Click
                    dlpDate(25).SetFocus
                    Exit Function
                End If
            End If
        End If
        If IsDate(dlpDate(24).Text) And IsDate(dlpDate(25).Text) Then
            If dlpDate(25).Enabled Then
                If CVDate(dlpDate(24).Text) > CVDate(dlpDate(25).Text) Then
                    Msg = "The NGS Start Date cannot be greater than the NGS End Date. " & Chr(10)
                    Msg = Msg & "If the employee was transferred to a new NGS Sub Group and NGS has " & Chr(10)
                    Msg = Msg & "already been notified of the change, remove the NGS End Date."
                    MsgBox Msg
                    'dlpDate(25).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    'Ticket #22584 Franks 10/03/2012 "   Do not allow the user to change the UNION
    glbMsgCustomVal = 0
    If Len(SavOrg) > 0 And NewHireForms.count = 0 Then
        If Not (SavOrg = clpCode(2).Text) Then
            '''Ticket #23903 Franks 06/11/2013
            '''check if it has been exported, if not yet then user can change Union here
            ''If IsEmpExported(glbLEE_ID) Then
            '    MsgBox lStr("Union") & " code cannot be changed on this screen." & Chr(10) & "Use the Transfer Out/IN screens to change an employee's Union Code."
            '    clpCode(2).SetFocus
            '    Exit Function
            ''End If
            'Ticket #23903 Franks 06/19/2013 - begin
            If glbDeptAllRight Then
                'users whose has a Department/Security Matrix = “ALL”, Pop up a message box to enter a password. Password is “petman”
                MsgBox lStr("Union") & " code was changed on this screen. You need to enter the correct password."
                glbAccessPswd = False
                frmAccessPswd.Show 1
                If glbAccessPswd = False Then   'Access Denied
                    clpCode(2).SetFocus
                    Exit Function
                End If
            Else
                ''''Ticket #23903 Franks 06/11/2013
                ''''check if it has been exported, if not yet then user can change Union here
                '''If IsEmpExported(glbLEE_ID) Then
                ''    MsgBox lStr("Union") & " code cannot be changed on this screen." & Chr(10) & "Use the Transfer Out/IN screens to change an employee's Union Code."
                ''    clpCode(2).SetFocus
                ''    Exit Function
                '''End If
                frmMsgDialog.lblMsg = "Is this change of Union due to a data correction or is the employee physically moving to another Union?"
                frmMsgDialog.OKButton.Caption = "Correct" ' If Correct is clicked, do not make a change to the union. Pop up another message saying “Please contact Corporate HRIS Manager to have this change made.”
                frmMsgDialog.CancelButton.Caption = "Moving" ' If Moving is clicked, do not make the Union change. Close the Status/Dates screen and open the Transfer Out screen
                frmMsgDialog.Show 1
                If glbMsgCustomVal = 1 Then 'Correct
                    'clpCode(2).Text = SavOrg 'do not make a change
                    MsgBox "Please contact Corporate HRIS Manager to have this change made."
                    Call cmdCancel_Click 'Ticket #23903 Franks 06/20/2013
                End If
                If glbMsgCustomVal = 2 Then 'Moving
                    'clpCode(2).Text = SavOrg 'do not make a change
                    Call cmdCancel_Click 'Ticket #23903 Franks 06/20/2013
                End If
            End If
            'Ticket #23903 Franks 06/19/2013 - end
        End If
    End If
End If
'------------WFC -------- end

If glbCompSerial = "S/N - 2376W" Then
    If clpPT.Text = "TERM" Or clpPT.Text = "PERM" Then
        If Len(dlpDate(3).Text) = 0 Or IsDate(dlpDate(3).Text) = False Then
            MsgBox lStr("User Defined date is mandatory")
            dlpDate(3).SetFocus
            Exit Function
        End If
    End If
End If

If (glbCompSerial = "S/N - 2409W") Then 'Ticket #30066 Franks - Skylark Children
    If Len(dlpDate(6)) = 0 Then
        tabDates.Tabs(1).selected = True
        MsgBox lStr("Seniority Date is a required field")
        dlpDate(6).SetFocus
        Exit Function
    End If
End If

If (glbCompSerial = "S/N - 2385W") Then ' Conservation Halton 'Ticket #13063
    If Len(dlpDate(6)) = 0 Then
        MsgBox lStr("Seniority Date is a required field")
        dlpDate(6).SetFocus
        Exit Function
    End If
    If clpPT.Text = "FT" Then
        If Len(dlpDate(4)) = 0 Then
            MsgBox lStr("Union Date is a required field")
            dlpDate(4).SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(2).Text) = 0 Then
            MsgBox lblUnion.Caption & (" is a required field")
            clpCode(2).SetFocus
            Exit Function
    End If
    If Len(Trim(clpBGroup.Text)) = 0 Then 'Conservation Halton Ticket #14402
        If clpCode(2).Text = "AD" Then 'Or clpCode(2).Text = "JS" Or clpCode(2).Text = "PB" Then 'Full Time Employees
            MsgBox lStr("Benefit Group is a required field for 'AD'")
            clpBGroup.SetFocus
            Exit Function
        End If
    End If

End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    If Len(clpCode(2).Text) <= 0 Then
        MsgBox lblUnion.Caption & " must be entered"
        clpCode(2).SetFocus
        Exit Function
    End If
    
    'Ticket #22178
    If Len(Trim(clpSalDist.Text)) = 0 Then
        MsgBox lblSalDist.Caption & " is a required field"
        clpSalDist.SetFocus
        Exit Function
    End If
    
    If Len(dlpDate(6)) = 0 Then
        MsgBox lStr("Seniority Date is a required field")
        dlpDate(6).SetFocus
        Exit Function
    End If
    
    'Ticket #20600 Franks 09/02/2011 - begin
    'If "Supervisor" is entered, OMERS Date must be entered too
    'If Len(clpCode(8).Text) > 0 Then
    'Ticket #22262 Franks 085/02/2012 - pension code field should be Salary Distribution not Supervisor Code
    If Len(clpSalDist.Text) > 0 Then
        'Ticket #22262 Franks 085/02/2012
        '1. If Original Hire date is greater than and equal to 01/01/2012, Employment Status is “A or ACEX” and Employment Category is “FT, FT65 or FTEX” = Pension code is Z
        '2. If Original Hire date is greater than and equal to 01/01/2012, Employment Status is “A or ACEX” and Employment Category is “FTC, PTC, PT, ST, CL, CON or UPT” = Pension code is X
        If IsDate(dlpDate(7).Text) Then
            If CVDate(dlpDate(7).Text) >= CVDate("01/01/2012") Then
                If clpCode(1).Text = "A" Or clpCode(1).Text = "ACEX" Then
                    If clpPT.Text = "FT" Or clpPT.Text = "FT65" Or clpPT.Text = "FTEX" Then
                        If Not clpSalDist.Text = "Z" Then
                            Msg = lStr("Salary Distribution") & " must be 'Z' if Employment Status is" & Chr(10) & "A or ACEX and Employment Category is FT, FT65 or FTEX."
                            Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to keep " & lStr("Salary Distribution") & " as '" & clpSalDist.Text & "' "
                            'MsgBox Msg
                            a% = MsgBox(Msg, 36, "Confirm ")
                            If a% <> 6 Then
                                clpSalDist.SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                    If clpPT.Text = "FTC" Or clpPT.Text = "PTC" Or clpPT.Text = "PT" Or clpPT.Text = "ST" Or clpPT.Text = "CL" Or clpPT.Text = "CON" Or clpPT.Text = "UPT" Then
                        If Not clpSalDist.Text = "X" Then
                            Msg = lStr("Salary Distribution") & " must be 'X' if Employment Status is A or ACEX and " & Chr(10) & "Employment Category is FTC, PTC, PT, ST, CL, CON or UPT."
                            Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to keep " & lStr("Salary Distribution") & " as '" & clpSalDist.Text & "' "
                            'MsgBox Msg
                            a% = MsgBox(Msg, 36, "Confirm ")
                            If a% <> 6 Then
                                clpSalDist.SetFocus
                                Exit Function
                            End If
                            'MsgBox lStr("Salary Distribution") & " must be 'X' if Employment Status is A or ACEX and " & Chr(10) & "Employment Category is FTC, PTC, PT, ST, CL, CON or UPT."
                        End If
                    End If
                End If
            End If
        End If
        
        If Len(dlpDate(2).Text) = 0 Then
            MsgBox lStr("OMERS Date") & " is required if " & lStr("Salary Distribution") & " is entered."
            dlpDate(2).SetFocus
            Exit Function
        End If
    End If
    If Len(clpVadim1.Text) > 0 Then
        If Len(dlpDate(3).Text) = 0 Then
            MsgBox lStr("User Defined") & " is required if " & lStr("Vadim Field 1") & " is entered."
            dlpDate(3).SetFocus
            Exit Function
        End If
    End If
    'Ticket #20600 Franks 09/02/2011 - end
    
    'Ticket #22491 Franks 08/30/2012 - begin - for Samuel
    If Not oEmail = txtEmail.Text Then
        If CheckDupEmpEmail(glbLEE_ID, txtEmail.Text) Then
            Load frmMsgBox
            frmMsgBox.cmdCancel.Caption = "No"
            frmMsgBox.cmdOk.Caption = "Yes"
            frmMsgBox.Caption = "Duplicate Email Address found "
            Msg$ = ""
            If Len(Trim(fDupEmail_Act)) > 0 Then
                frmMsgBox.txtLongMsg = fDupEmail_Act & vbNewLine & vbNewLine & fDupEmail_Term
            Else
                frmMsgBox.txtLongMsg = fDupEmail_Term
            End If
            Msg$ = "Are you sure you wish to accept it?"
            Msg$ = Msg$ & Chr(10) & "Press Yes to accept or No to edit"
            frmMsgBox.lblQuestion = Msg$
            frmMsgBox.Show 1
            If glbMsgBoxResult = vbCancel Then
                txtEmail.SetFocus
                Exit Function
            End If
        End If
    End If
    'Ticket #22491 Franks 08/30/2012 - end
End If

If glbCompSerial = "S/N - 2335W" Then 'Mitchell Plastics Ticket #21866 Franks 04/05/2012
    If Len(clpCode(2).Text) = 0 Then
            MsgBox lblUnion.Caption & (" is a required field")
            clpCode(2).SetFocus
            Exit Function
    End If
End If
If glbCompSerial = "S/N - 2451W" Then 'Decor Ticket #23848
    If Len(clpCode(2).Text) = 0 Then
            MsgBox lblUnion.Caption & (" is a required field")
            clpCode(2).SetFocus
            Exit Function
    End If
End If

If glbWFC Then 'Ticket #25248 Franks 03/24/2014
    If lblWFCMsg.Visible Then
        If Len(dlpDate(29).Text) = 0 Then 'Ticket #27609 Franks 10/07/2015
            MsgBox "401k Eligibility Date is missing"
            tabDates.Tabs(3).selected = True
            Call tabDates_Click
            dlpDate(29).SetFocus
            Exit Function
        End If
    End If
    
    If clpCode(1).Text = "SALC" Then
        If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
            If Len(dlpDate(1).Text) = 0 Then
                MsgBox lblLDay.Caption & " is mandatory for 'NONE' or 'EXEC' employees" & Chr(10) & "if Employment Status is 'SALC'"
                tabDates.Tabs(1).selected = True
                Call tabDates_Click
                dlpDate(1).SetFocus
                Exit Function
            End If
        End If
        'Ticket #25352 Franks 04/16/2014 - begin
        '"   If Status = SALC, they can't change the Last Day.
        If IsDate(dlpDate(1).Text) And IsDate(OLDAY) Then
            If Not (CVDate((dlpDate(1).Text)) = CVDate(OLDAY)) Then
                MsgBox lblLDay.Caption & " cannot be changed if Employment Status is 'SALC'"
                tabDates.Tabs(1).selected = True
                Call tabDates_Click
                dlpDate(1).SetFocus
                Exit Function
            End If
        End If
        'Ticket #25352 Franks 04/16/2014 - end
    End If
End If

'Ticket #25469 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    If Len(clpCode(2).Text) = 0 Then
        MsgBox lblUnion.Caption & (" is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If

    'They don't want Email Address to be Mandatory
    'If Len(txtEmail.Text) = 0 Then
    '    MsgBox "Email address is required field"
    '    txtEmail.SetFocus
    '    Exit Function
    'End If
End If

'Ticket #26008 - Surrey Place - Other Dates 1 - make it mandatory
If glbCompSerial = "S/N - 2347W" Then
    If Len(dlpDate(24).Text) < 1 Then
        MsgBox lStr(lbOtherDate(0).Caption) & " is a required field"
        tabDates.Tabs(3).selected = True 'Ticket #26008 Franks 10/30/2014
        Call tabDates_Click
        dlpDate(24).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDate(24).Text) Then
        MsgBox lStr(lbOtherDate(0).Caption) & " is not a valid date"
        tabDates.Tabs(3).selected = True 'Ticket #26008 Franks 10/30/2014
        Call tabDates_Click
        dlpDate(24).SetFocus
        Exit Function
    End If
End If

''If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
''    If Len(dlpDate(24).Text) < 1 Then
''        MsgBox lStr(lbOtherDate(0).Caption) & " is a required field"
''        tabDates.Tabs(3).selected = True
''        Call tabDates_Click
''        If dlpDate(24).Enabled Then dlpDate(24).SetFocus
''        Exit Function
''    End If
''End If

''Ticket #29230 - Daily Entitlement - If these values change then there must be respective Effective Date so daily accrual can be re-computed from Effective Date onwards
''Employment Status, Category and Union
If glbCompEntVacDaily Then
    If SavEmp <> clpCode(1).Text Or SavPT <> clpPT.Text Or SavOrg <> clpCode(2).Text Then
        If SavEmp <> clpCode(1).Text And (Not IsDate(dlpDate(15).Text) Or Len(dlpDate(15).Text) = 0) Then
            MsgBox "Invalid From Date for Employment Status. For Daily Accrual re-computation 'From Date' is required when Employment Status changes."
            dlpDate(15).SetFocus
            Exit Function
        End If
        If SavPT <> clpPT.Text And (Not IsDate(dlpDate(34).Text) Or Len(dlpDate(34).Text) = 0) Then
            MsgBox "Invalid " & lStr("Category") & " Effective Date. For Daily Accrual re-computation '" & lStr("Category") & " Effective Date' is required when " & lStr("Category") & " changes."
            dlpDate(34).SetFocus
            Exit Function
        End If
        If SavOrg <> clpCode(2).Text And (Not IsDate(dlpDate(36).Text) Or Len(dlpDate(36).Text) = 0) Then
            MsgBox "Invalid " & lStr("Union") & " Effective Date. For Daily Accrual re-computation '" & lStr("Union") & " Effective Date' is required when " & lStr("Union") & " changes."
            dlpDate(36).SetFocus
            Exit Function
        End If
    End If
End If

chkEStats = True

End Function

Private Function chkLOACode(xCode) As Boolean
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
    chkLOACode = False
    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_NAME = 'EDEM' AND TB_KEY = '" & xCode & "' AND NOT (TB_USR3 = 0) "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTABL.EOF Then
        chkLOACode = True
    End If
    rsTABL.Close
End Function

Private Sub chkSpouse_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Public Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err



''' Sam add July 2002 * Remove ADO
If glbtermopen Then
    rsDATA2.CancelUpdate
End If
rsDATA.CancelUpdate
 
 
fglbNew = False
Call SET_UP_MODE
Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEESTATS" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdModify_Click()
'Dim rsTB As Recordset
Dim strDOHDate, strFMLADate, strDOHYear
Dim strFMLAYear, strYear, strDecade, strCentury
Dim x%

On Error GoTo Mod_Err

fglHredsem = dlpDate(1).Text

SavEmp = clpCode(1).Text
savLOA = chkLeave
SavOrg = clpCode(2).Text
SavPT = clpPT.Text
OEmptype = txtEmpType
SaveBGroup = clpBGroup.Text
SavDOH = dlpDate(7).Text
For x% = 1 To 7
    ODate(x%) = dlpDate(x%).Text
Next x%

ODate(8) = dlpDate(0).Text

'laura nov 11, 1997
For x% = 8 To 12
    If (glbCompSerial = "S/N - 2375W") And (x% = 10) And (NewHireForms.count > 0) Then 'City of Timmins
        'Do not save Normal Retirement Date into ODate because it's then not passed to Vadim
    Else
        ODate(x% + 1) = dlpDate(x%).Text
    End If
Next x%
ODeptEDate = dlpDate(13).Text
ODivEdate = dlpDate(14).Text

oFDate = dlpDate(15).Text
OTDate = dlpDate(16).Text
OPTEDate = dlpDate(34).Text
'Ticket #29230 - Daily Vacation Entitlement
OORGEDate = dlpDate(36).Text

OINTEL = txtIPHONE
'OLANG1 = clpCode(3).Text
'OLANG2 = clpCode(4).Text
oHireCode = clpCode(6).Text
OOMERS = dlpDate(2).Text
OLDAY = dlpDate(1).Text
oEmail = txtEmail
OBenGrp = clpBGroup.Text
oSalDist = clpSalDist
OWITHSPOUSE = chkSpouse
OEXPYEAR = txtExpYear
oUSER_TEXT1 = txtUserText1.Text
oUSER_TEXT2 = txtUserText2.Text
oUSER_NUM1 = txtUserNum1.Text
oUSER_NUM2 = txtUserNum2.Text '""

oPENSIONDATE1 = dlpDate(18).Text
oPENSIONDATE2 = dlpDate(19).Text
oPENSIONDATE3 = dlpDate(20).Text
oPENSIONDATE4 = dlpDate(21).Text
oPENSIONDATE5 = dlpDate(22).Text
oPENSIONDATE6 = dlpDate(23).Text

oOTHERDATE1 = dlpDate(24).Text
oOTHERDATE2 = dlpDate(25).Text
oOTHERDATE3 = dlpDate(26).Text
oOTHERDATE4 = dlpDate(27).Text
oOTHERDATE5 = dlpDate(28).Text
oOTHERDATE6 = dlpDate(29).Text
oOTHERDATE7 = dlpDate(30).Text
oOTHERDATE8 = dlpDate(31).Text
oOTHERDATE9 = dlpDate(32).Text
oOTHERDATE10 = dlpDate(33).Text
If glbtermopen Then
    OTermDate = dlpTermDate
End If
If glbWFC Then 'Ticket #19266
    OVadim11 = clpVadim1.Text
    OVadim21 = clpVadim2.Text
End If
If clpCode(8).Visible Then
    oSuperCode = clpCode(8).Text
    OVadim11 = clpVadim1.Text
End If
If medVacPPct.Visible Then 'Ticket #22710
    OVACPC = medVacPPct
End If

'Ticket #24996 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    OSection = clpCode(4).Text
End If

Call FMLA  'jaddy oct 4,99

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREMP", "Modify")
Call RollBack '23June99 - js

End Sub

'Private Sub cmdModify_GotFocus()
 '   Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()
Dim DtTm As Variant, rc As Integer, xEMP, xORG, xPT
Dim xTDate, xFDate, xtmpdate, xPTDate, xOrgDate
Dim x%, Y%, SQLQ, Msg
Dim xWOldFTPT
Dim rsHRJOB As New ADODB.Recordset
Dim xFlag1 As Boolean 'Ticket #20169
Dim xFlag2 As Boolean 'Ticket #20169
Dim xFlag3 As Boolean 'Ticket #20169

'Release 8.1
Dim rsBenCode As New ADODB.Recordset

DtTm = Now
If Not chkEStats() Then Exit Sub

If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Ticket #14752
    If NewHireForms.count > 0 Then  'For New Hires Only
        Call Popup_JobCode
    Else
        'Comment by Frank Ticket #14752, this Job code pop up function only works for new hire
        ''Check if Position Code exists
        'SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <>0 AND JH_EMPNBR = " & glbLEE_ID
        'rsHRJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'If rsHRJob.EOF Then
        '    rsHRJob.Close
        '    Set rsHRJob = Nothing
        '    Call Popup_JobCode
        'End If
    End If
End If

rsDATA.Requery

Call UpdUStats(Me)

If Not glbtermopen Then
    If Not glbLinamar Then
        If Not UpdEntOut() Then Exit Sub
    End If
    
    If (glbCompSerial = "S/N - 2388W") Then 'DNSSAB Ticket #14333
        If Len(Trim(SavPT)) > 0 And SavPT <> clpPT.Text Then
            If SavPT = "FT" And clpPT.Text = "PT" Then
                If Len(Trim(clpBGroup.Text)) > 0 Then
                    clpBGroup.Text = ""
                End If
            End If
        End If
    End If
    
    Call UpdEICodeForGraniteClub
    
    'Ticket #29230 - Daily Vacation Entitlement
    If SavEmp <> clpCode(1).Text Or SavOrg <> clpCode(2).Text Or SavPT <> clpPT.Text Or oFDate <> dlpDate(15).Text Or OTDate <> dlpDate(16).Text Or OPTEDate <> dlpDate(34).Text Or OORGEDate <> dlpDate(36).Text Then   'Hemu EMPHIS
    'If SavEmp <> clpCode(1).Text Or SavOrg <> clpCode(2).Text Or SavPT <> clpPT.Text Or oFDate <> dlpDate(15).Text Or OTDate <> dlpDate(16).Text Or OPTEDate <> dlpDate(34).Text Then   'Hemu EMPHIS
        xEMP = ""
        xORG = ""
        xPT = ""
        xTDate = ""
        xFDate = ""
        xPTDate = ""
        If SavEmp <> clpCode(1).Text Then
            'If Len(clpCode(1).Text) > 0 Then xEMP = clpCode(1).Text Else xEMP = "*"

            'Hemu
            If Len(clpCode(1).Text) > 0 Then
                xEMP = clpCode(1).Text
                If Not IsDate(dlpDate(15).Text) And Not IsDate(dlpDate(16)) Then
                    xFDate = Date
                Else
                    xFDate = dlpDate(15)
                    xTDate = dlpDate(16)
                End If
            Else
                xEMP = "*"
            End If
        Else
            If oFDate <> dlpDate(15) Or OTDate <> dlpDate(16) Then
                'Go back to the HREMPHIS table and change that record
                SQLQ = "UPDATE HREMPHIS SET EE_CHGDATE =" & Date_SQL(dlpDate(15).Text) & ", EE_TODATE = " & Date_SQL(dlpDate(16).Text)
                SQLQ = SQLQ & " WHERE EE_EMPNBR = " & glbLEE_ID & " AND EE_NEWSTAT = '" & clpCode(1).Text & "'"
                If IsDate(oFDate) And IsDate(OTDate) Then
                    SQLQ = SQLQ & " AND EE_CHGDATE = " & Date_SQL(oFDate) & " AND EE_TODATE = " & Date_SQL(OTDate)
                Else
                    If Not IsDate(oFDate) And Not IsDate(OTDate) Then
                        SQLQ = SQLQ & " AND EE_CHGDATE IS NULL AND EE_TODATE IS NULL"
                    ElseIf Not IsDate(oFDate) Then
                        SQLQ = SQLQ & " AND EE_CHGDATE IS NULL AND EE_TODATE = " & Date_SQL(OTDate)
                    ElseIf Not IsDate(OTDate) Then
                        SQLQ = SQLQ & " AND EE_CHGDATE = " & Date_SQL(oFDate) & " AND EE_TODATE IS NULL"
                    End If
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
            'Hemu
        End If
        
        'Ticket #29230 - Daily Vacation Entitlement
        If SavOrg <> clpCode(2).Text Or OORGEDate <> dlpDate(36).Text Then
        'If SavOrg <> clpCode(2).Text Then
            If Len(clpCode(2).Text) > 0 Then xORG = clpCode(2).Text Else xORG = "*"
        End If
        If SavPT <> clpPT.Text Or OPTEDate <> dlpDate(34).Text Then
            If Len(clpPT.Text) > 0 Then xPT = clpPT.Text Else xPT = "*"
        End If
        
        If xEMP <> "*" And Len(xEMP) > 0 Then
            'If Not EmpHisCalc(1, glbLEE_ID, "", "", xEMP, xPT, xORG, "", "", Date, , , xfdate, xtdate) Then MsgBox "EMPHIS Error"
            If NewHireForms.count > 0 Then 'Ticket #23837 Franks 05/28/2013
                If Not EmpHisCalc(1, glbLEE_ID, "", "", xEMP, "", "", "", "", dlpDate(7).Text, , , xFDate, xTDate) Then MsgBox "EMPHIS Error"
            Else
                If Not EmpHisCalc(1, glbLEE_ID, "", "", xEMP, "", "", "", "", Date, , , xFDate, xTDate) Then MsgBox "EMPHIS Error"
            End If
        End If
        'Else
            'If Not EmpHisCalc(1, glbLEE_ID, "", "", xEMP, xPT, xORG, "", "", Date) Then MsgBox "EMPHIS Error"
            'If Not EmpHisCalc(1, glbLEE_ID, "", "", "", xPT, xORG, "", "", Date) Then MsgBox "EMPHIS Error"
        'End If
        
        'Ticket #18035 - Jerry asked to add ED_PT as CATEGORY in the Employee History table
        If xPT <> "*" And Len(xPT) > 0 Then
            If NewHireForms.count > 0 Then 'Ticket #23837 Franks 05/28/2013
                If Not IsDate(dlpDate(34).Text) Then xPTDate = dlpDate(7).Text Else xPTDate = dlpDate(34).Text
            Else
                If Not IsDate(dlpDate(34).Text) Then xPTDate = Date Else xPTDate = dlpDate(34).Text
            End If
            If Not EmpHisCalc(1, glbLEE_ID, "", "", "", xPT, "", "", "", xPTDate) Then MsgBox "EMPHIS Error"
        End If
        
        'Ticket #22584 Franks 10/03/2012 - Jerry asked to add Union in the Employee History table
        If xORG <> "*" And Len(xORG) > 0 Then
            If NewHireForms.count > 0 Then 'Ticket #23837 Franks 05/28/2013
                'Ticket #29230 - Daily Vacation Entitlement
                If Not IsDate(dlpDate(36).Text) Then xOrgDate = dlpDate(7).Text Else xOrgDate = dlpDate(36).Text
                If Not EmpHisCalc(1, glbLEE_ID, "", "", "", "", xORG, "", "", xOrgDate) Then MsgBox "EMPHIS Error"
                'If Not EmpHisCalc(1, glbLEE_ID, "", "", "", "", xORG, "", "", dlpDate(7).Text) Then MsgBox "EMPHIS Error"
            Else
                'Ticket #29230 - Daily Vacation Entitlement
                If Not IsDate(dlpDate(36).Text) Then xOrgDate = Date Else xOrgDate = dlpDate(36).Text
                If Not EmpHisCalc(1, glbLEE_ID, "", "", "", "", xORG, "", "", xOrgDate) Then MsgBox "EMPHIS Error"
            End If
        End If
        
        'Ticket #19785 - Update employee's Current Position's multiposition fields with default values for employees
        'with non multi position setting (ED_SECTION <> "Y").
        If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford
            If rsDATA("ED_SECTION") <> "Y" Then
                Call UpdateOxfordCurrentPosition
            End If
        End If
        
        'Ticket #27899 - Also for WDGPHU as above
        If glbCompSerial = "S/N - 2411W" Then   'For WDGPHU
            If rsDATA("ED_ORGT1") <> "YES" Then
                Call UpdateOxfordCurrentPosition
            End If
        End If
        
        'Kerry's Place - Ticket #24692 - Update matching Current Positions fields
        If glbCompSerial = "S/N - 2433W" Then
            If SavOrg <> clpCode(2).Text Then
                Call UpdateCurrentPosition_KerrysPlace
            End If
        End If
        
    End If
    
    ''Ticket #24179 Franks 085/06/2013
    ''If SaveBGroup <> clpBGroup.Text Then
    ''    NewBGroup = Trim(clpBGroup.Text)
    ''    If NewHireForms.count > 0 Then 'Ticket #23837 Franks 05/28/2013
    ''        If Not EmpHisCalc(0, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(7).Text, "BENEGROUP", NewBGroup) Then MsgBox "EMPHIS Error"
    ''    Else
    ''        If Not EmpHisCalc(0, glbLEE_ID, "", "", "", "", "", "", "", Date, "BENEGROUP", NewBGroup) Then MsgBox "EMPHIS Error"
    ''    End If
    ''    'Ticket #20169 Franks 04/21/2011 check if there is benefit group setup for this code,
    ''    xFlag1 = BenGroupExist(SaveBGroup)
    ''    xFlag2 = BenGroupExist(NewBGroup)
    ''    xFlag3 = False
    ''    If xFlag1 Or xFlag2 Then 'Benefit Group setup for Old Group or New Group
    ''            Msg = "Do you want add/update the Employee's Benefits "
    ''            Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
    ''            If MsgBox(Msg, 36, "INFO:HR") = 6 Then
    ''                xFlag3 = True
    ''                Call UpdateBenefitGroup
    ''                DoEvents
    ''                If glbVadim Then
    ''                    frmBENGRLIST.dlpProcessDate = dlpDate(4)
    ''                End If
    ''                frmBENGRLIST.Show 1
    ''            End If
    ''    End If
    ''    'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
    ''    If Not xFlag3 Then
    ''        If Len(clpBGroup.Text) = 0 Then
    ''            SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & lblEEID
    ''            gdbAdoIhr001.Execute SQLQ
    ''        End If
    ''    End If
    ''    SaveBGroup = clpBGroup.Text
    ''End If
    
    'If glbCompSerial = "S/N - 2172W" Then
    'Ticket #19782 Franks 02/02/2011 for Frontenac
    If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
        Call UpdateGPMainBenDed
    End If
    
    If glbWFC Then
        glbChgTermDate = ""
        glbChgTermReason = ""
        'Ticket #19086 Frank on 09/09/2010, remove this logic
        'If rsDATA("ED_SECTION") = "TILB" And Not IsNull(rsDATA("ED_PAYROLL_ID")) Then ' TILBURY
        'Dim xWOldStatus
        '    xWOldStatus = IIf(IsNull(rsDATA("ED_EMP")), "", rsDATA("ED_EMP"))
        '    If xWOldStatus = "PROB" And Len(rsDATA("ED_PAYROLL_ID")) > 0 Then
        '
        '        Screen.MousePointer = DEFAULT
        '        'For TILBURY, if ED_EMP was changed from "PROB" to other, then
        '        'Create terminated and new hire records in HRAUDIT
        '        If xWOldStatus <> clpCode(1) Then
        '            If Len(xWOldStatus) > 0 And Len(clpCode(1).Text) > 0 Then
        '                frmSoroc.Show 1
        '            End If
        '        End If
        '        Screen.MousePointer = HOURGLASS
        '    End If
        'End If
        'Greensboro
        If rsDATA("ED_SECTION") = "GREN" And Not IsNull(rsDATA("ED_PAYROLL_ID")) Then
        'Dim xWOldFTPT
            xWOldFTPT = IIf(IsNull(rsDATA("ED_PT")), "", rsDATA("ED_PT"))
            glbEESection = rsDATA("ED_SECTION")
            If xWOldFTPT <> clpPT Then
                If clpPT = "TR" And xWOldFTPT = "FT" Then
                    Screen.MousePointer = DEFAULT
                    If Len(xWOldFTPT) > 0 And Len(clpPT.Text) > 0 Then
                        frmSoroc.Show 1
                    End If
                    Screen.MousePointer = HOURGLASS
                End If
                If xWOldFTPT = "TR" And clpPT = "FT" Then
                    Screen.MousePointer = DEFAULT
                    If Len(xWOldFTPT) > 0 And Len(clpPT.Text) > 0 Then
                        frmMsgTerm.Show 1
                    End If
                    Screen.MousePointer = HOURGLASS
                End If
            End If
        End If
        'NGS Transaction - begin - Ticket #19266
        If NewHireForms.count = 0 Then
            Call WFC_NGS_Trans
        End If
        'NGS Transaction - end
        'Ticket #20136 - Franks 04/11/2011
        If NewHireForms.count > 0 Then
            Call chkWFCOtherDate6("Upt")
        End If
    End If
    
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        glbChgTermDate = ""
        glbChgTermReason = ""
        glbChgNewEmpnbr = ""
        glbEESection = ""
        If Not IsNull(rsDATA("ED_PAYROLL_ID")) Then
        'Dim xWOldFTPT
            xWOldFTPT = IIf(IsNull(rsDATA("ED_PT")), "", rsDATA("ED_PT"))
            If xWOldFTPT <> clpPT Then
                If (xWOldFTPT = "FT" Or xWOldFTPT = "PT") And (clpPT = "FT" Or clpPT = "PT") Then
                    Screen.MousePointer = DEFAULT
                    If Len(xWOldFTPT) > 0 And Len(clpPT.Text) > 0 Then
                        frmMsgTerm.Show 1
                    End If
                    Screen.MousePointer = HOURGLASS
                End If
            End If
        End If
    End If
    
    If Not AUDITSTAT() Then MsgBox "ERROR : AUDIT FILE"
    If Not AUDITSTAT2() Then MsgBox "ERROR : AUDIT2 FILE"
    If glbWFC Then 'Manulife Audit
        Call AUDIT_MANULIFE_TRANS
    End If
    If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21518 Franks 05/03/2012
        Call AUDIT_GWL_TRANS
    End If
    
End If

'Pass Terminated Employee changes to Vadim
If glbVadim And glbtermopen Then
    Call Pass_TermEmp_Change_Vadim
End If

If Not glbLinamar Then
    Call UPDStatusLOG
End If

If Len(txtExpYear) > 0 Then
    rsDATA("ED_EXPYEAR") = Val(txtExpYear)
End If

If glbCompSerial = "S/N - 2296W" Then  'For Essex Library
    If OOMERS <> dlpDate(2).Text Then
        rsDATA("ED_PENSION") = "1"
    End If
End If

'Added by Bryan 30/Sep/05 Ticket#9431
If OOMERS <> dlpDate(2).Text And glbCompSerial = "S/N - 2376W" Then 'Assemby First Nations
    Dim strSQL As String
    If IsDate(dlpDate(2).Text) Then
        strSQL = "INSERT INTO HR_FOLLOW_UP (EF_COMPNO, EF_EMPNBR, EF_FDATE, EF_FREAS_TABL, EF_FREAS, EF_LDATE, EF_LTIME, EF_LUSER) "
        strSQL = strSQL & "VALUES ('001', " & glbLEE_ID & ", " & Date_SQL(dlpDate(2).Text) & ", 'FURE', 'PEN', " & Date_SQL(Date) & ", '" & Time$ & "', '" & glbUserID & "')"
        
        gdbAdoIhr001.Execute strSQL
    End If
End If

'Added by Bryan 30/Sep/05 Ticket#9431
'modified by Bryan 13/Oct/05 Ticket#9507
If glbCompSerial = "S/N - 2376W" And Len(dlpDate(1).Text) > 0 And NewHireForms.count > 0 Then  'Assemby First Nations
    strSQL = "INSERT INTO HR_FOLLOW_UP (EF_COMPNO, EF_EMPNBR, EF_FDATE, EF_FREAS_TABL, EF_FREAS, EF_LDATE, EF_LTIME, EF_LUSER) "
    strSQL = strSQL & "VALUES ('001', " & glbLEE_ID & ", " & Date_SQL(dlpDate(1).Text) & ", 'FURE', 'TE', " & Date_SQL(Date) & ", '" & Time$ & "', '" & glbUserID & "')"
    gdbAdoIhr001.Execute strSQL
End If

If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    If NewHireForms.count > 0 Then 'New Hire only
        If clpCode(1) = "RFT" Or clpCode(1) = "RPT" Then
            rsDATA("ED_UIC") = "1"
        Else
            rsDATA("ED_UIC") = "2"
        End If
        rsDATA("ED_WCBCODE") = "W"
    End If
End If

'Simcoe Muskoka District Health Unit Ticket #12302
If glbCompSerial = "S/N - 2228W" Then
    If NewHireForms.count > 0 Then 'New Hire only
        If Len(dlpDate(3).Text) = 0 Then 'if user defined date is blank
            dlpDate(3).Text = dlpDate(6).Text 'copy seniority to user defined
        End If
    End If
End If
        
If glbCompSerial = "S/N - 2205W" Then 'Crown Investment Corp. - Ticket #14084
    'New Hire Only for FT employee only
    'if Seniority, Union or User Defined date is blank then copy Date of Hire
    If NewHireForms.count > 0 And clpPT.Text = "FT" Then 'New Hire only
        If Len(dlpDate(6).Text) = 0 Then    'Seniority Date
            dlpDate(6).Text = dlpDate(7).Text
        End If
        If Len(dlpDate(4).Text) = 0 Then    'Union Date
            dlpDate(4).Text = dlpDate(7).Text
        End If
        If Len(dlpDate(3).Text) = 0 Then    'User Defined
            dlpDate(3).Text = dlpDate(7).Text
        End If
    End If
End If

'Bird Packaging 'Ticket #13701 On New Hire, ADP Data Control 1 = 0. Send to ADP
If glbCompSerial = "S/N - 2387W" Then
    If NewHireForms.count > 0 Then 'New Hire only
        Call ADP_Control(glbLEE_ID, "0")
    End If
End If


If glbCompSerial = "S/N - 2375W" Then  'City of Timmins
    If txtRPP.Text = "Nu" Then
        txtRPP.Text = ""
    End If
End If

If glbtermopen Then
    Call Set_Control2("U", rsDATA2)
    rsDATA2.Update
End If
    
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
    
Call Set_Control("U", Me, rsDAT_Other, True)
rsDAT_Other.Update
    
'Ticket #29230 - Daily Entitlement - If these values have changes then re-computed the Daily Accrual from Effective Date onwards
'Employment Status, Category and Union
If Not glbtermopen Then
    If glbCompEntVacDaily Then
        If SavEmp <> clpCode(1).Text Or SavPT <> clpPT.Text Or SavOrg <> clpCode(2).Text Or oFDate <> dlpDate(15).Text Or OPTEDate <> dlpDate(34).Text Or OORGEDate <> dlpDate(36).Text Then
            If (SavEmp <> clpCode(1).Text And IsDate(dlpDate(15).Text) And Len(dlpDate(15).Text) > 0) Or (oFDate <> dlpDate(15).Text And IsDate(dlpDate(15).Text) And Len(dlpDate(15).Text) > 0) Then
                Call Recompute_DailyAccrualFile(glbLEE_ID, dlpDate(15).Text)
            End If
            If (SavPT <> clpPT.Text And IsDate(dlpDate(34).Text) And Len(dlpDate(34).Text) > 0) Or (OPTEDate <> dlpDate(34).Text And IsDate(dlpDate(34).Text) And Len(dlpDate(34).Text) > 0) Then
                Call Recompute_DailyAccrualFile(glbLEE_ID, dlpDate(34).Text)
            End If
            If (SavOrg <> clpCode(2).Text And IsDate(dlpDate(36).Text) And Len(dlpDate(36).Text) > 0) Or (OORGEDate <> dlpDate(36).Text And IsDate(dlpDate(36).Text) And Len(dlpDate(36).Text) > 0) Then
                Call Recompute_DailyAccrualFile(glbLEE_ID, dlpDate(36).Text)
            End If
        End If
    End If
End If
    
'Ticket #24179 Franks 08/06/2013 - move this function down since it need the current ED_BENEFIT_GROUP
'to do the benefit salary dependent calculate
If SaveBGroup <> clpBGroup.Text Then
    NewBGroup = Trim(clpBGroup.Text)
    If NewHireForms.count > 0 Then 'Ticket #23837 Franks 05/28/2013
        If Not EmpHisCalc(0, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(7).Text, "BENEGROUP", NewBGroup) Then MsgBox "EMPHIS Error"
    Else
        If Not EmpHisCalc(0, glbLEE_ID, "", "", "", "", "", "", "", Date, "BENEGROUP", NewBGroup) Then MsgBox "EMPHIS Error"
    End If
    'Ticket #20169 Franks 04/21/2011 check if there is benefit group setup for this code,
    xFlag1 = BenGroupExist(SaveBGroup)
    xFlag2 = BenGroupExist(NewBGroup)
    xFlag3 = False
    If xFlag1 Or xFlag2 Then 'Benefit Group setup for Old Group or New Group
        If IsWFCUSBenEmp(glbLEE_ID) Then 'US Benefit Ticket #23247 Franks 03/12/2014
            'Ticket #23247 Franks 03/12/2014 - begin
            If Len(NewBGroup) = 0 Then
                'remove Benefit Group then pop up Benefit list to let user decide
                'if they want to delete some benefits which are not in the new group but in the old group
                Msg = "Do you want add/update the Employee's Benefits "
                Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
                If MsgBox(Msg, 36, "info:HR") = 6 Then
                    xFlag3 = True
                    Call UpdateBenefitGroup
                    DoEvents
                    If glbVadim Then
                        frmBENGRLIST.dlpProcessDate = dlpDate(4)
                    End If
                    frmBENGRLIST.Show 1
                End If
            Else
                'add new benefit group,
                '"   The company paid benefits should automatically add just like they do in the new hire procedure.
                If glbMsgCustomVal = 8 And IsDate(dlpDate(24).Text) Then
                    DoEvents
                    'For US Benefit
                    '"   Changed this employee from PT to FT. The NGS End Date and the Benefit End Dates were not removed.
                    Call WFCUpdateBenefitEndDate(glbLEE_ID, Date, "RemoveEndDate")

                    'only for status change from PT to FT:
                    '"   Also, when the benefits are updated, the Effective date should be the x Waiting Period from the NGS Start Date and not the DOH or today's date. -
                    Call WFC_UptUSBenByEmp(glbLEE_ID, CVDate(dlpDate(24).Text), 0, "Y", "Y", , SaveBGroup, DateAdd("D", 1, CVDate(dlpDate(24).Text)), , CVDate(dlpDate(24).Text)) ', dlpDate(24).Text)
                Else
                    Call WFC_UptUSBenByEmp(glbLEE_ID, CVDate(dlpDate(7).Text), 0, "Y", "Y", , SaveBGroup, DateAdd("D", 1, Date)) ', dlpDate(7).Text)
                End If
            End If
            'Ticket #23247 Franks 03/12/2014 - end
        Else 'non US Benefits
            Msg = "Do you want add/update the Employee's Benefits "
            Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
            If MsgBox(Msg, 36, "info:HR") = 6 Then
                xFlag3 = True
                
                Call UpdateBenefitGroup
                DoEvents
                
                If glbVadim Then
                    frmBENGRLIST.dlpProcessDate = dlpDate(4)
                End If
                
                'Release 8.1 - Email will be sent on Benefit changes as well.
                glbBenAdded = ""
                glbBenChanged = ""
                glbBenDeleted = ""
                
                frmBENGRLIST.Show 1
            
                'Release 8.1 - Email will be sent on Benefit changes as well.
                If gsEMAIL_ONBENEFIT Then
                    If glbBenAdded <> "" Or glbBenDeleted <> "" Or glbBenChanged <> "" Then
                        'Send Email
                        MailBody = "The Benefit Update:" & vbCrLf & vbCrLf

                        MailBody = MailBody & "Employee #: " & lblEENUM.Caption & vbCrLf
                        MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                        
                        'Following Benefits were Added
                        If glbBenAdded <> "" Then
                            MailBody = MailBody & vbCrLf & "The following New Benefit(s) got Added : " & vbCrLf
                            
                            'Retrieve the Benefits added to get the Effective Date and Benefit Code Description
                            SQLQ = "SELECT BF_BCODE, BF_EDATE FROM HRBENFT WHERE BF_EMPNBR = " & lblEEID
                            SQLQ = SQLQ & " AND BF_BCODE IN ('" & Replace(glbBenAdded, ",", "','") & "')"
                            rsBenCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
                            Do While Not rsBenCode.EOF
                                'Mail Body
                                MailBody = MailBody & vbTab & " - " & GetTABLDesc("BNCD", rsBenCode("BF_BCODE")) & " with Effective Date: " & Format(rsBenCode("BF_EDATE"), "SHORT DATE") & vbCrLf
                                rsBenCode.MoveNext
                            Loop
                            rsBenCode.Close
                            Set rsBenCode = Nothing
                        End If
                                                
                        'Following Benefits were Updated
                        If glbBenChanged <> "" Then
                            MailBody = MailBody & vbCrLf & "The following Benefit(s) got Updated : " & vbCrLf
                            
                            'Retrieve the Benefits updated to get the Effective Date and Benefit Code Description
                            SQLQ = "SELECT BF_BCODE, BF_EDATE FROM HRBENFT WHERE BF_EMPNBR = " & lblEEID
                            SQLQ = SQLQ & " AND BF_BCODE IN ('" & Replace(glbBenChanged, ",", "','") & "')"
                            rsBenCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
                            Do While Not rsBenCode.EOF
                                'Mail Body
                                MailBody = MailBody & vbTab & " - " & GetTABLDesc("BNCD", rsBenCode("BF_BCODE")) & " with Effective Date: " & Format(rsBenCode("BF_EDATE"), "SHORT DATE") & vbCrLf
                                rsBenCode.MoveNext
                            Loop
                            rsBenCode.Close
                            Set rsBenCode = Nothing
                        End If
                        
                        'Following Benefits were Deleted
                        If glbBenDeleted <> "" Then
                            MailBody = MailBody & vbCrLf & "The following Benefit(s) got Deleted: " & vbCrLf
                            
                            'Retrieve the Benefits Deleted to get Benefit Code Description
                            SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_KEY IN ('" & Replace(glbBenDeleted, ",", "','") & "') "
                            SQLQ = SQLQ & " AND TB_NAME = 'BNCD'"
                            rsBenCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
                            Do While Not rsBenCode.EOF
                                'Mail Body
                                MailBody = MailBody & vbTab & " - " & rsBenCode("TB_DESC") & vbCrLf
                                rsBenCode.MoveNext
                            Loop
                            rsBenCode.Close
                            Set rsBenCode = Nothing
                        End If
                        
                        Call imgEmailBenefit_Click
                        
                        Screen.MousePointer = DEFAULT
                    End If
                End If
            End If
        End If
    End If
    'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
    If Not xFlag3 Then
        If Len(clpBGroup.Text) = 0 Then
            SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & lblEEID
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    
    'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
    If glbCompSerial = "S/N - 2290W" Then
        Call Update_Age65_LTD_Benefit_EndDate(glbLEE_ID, GetEmpData(glbLEE_ID, "ED_DOB"))
    End If
    
    SaveBGroup = clpBGroup.Text
Else 'SaveBGroup = clpBGroup.Text
    If glbWFC Then 'Ticket #24179 Franks 02/25/2014
        'US NGS employee status change to end benefits
        'talked this with Jerry, if the employee's Status was changed, then we will use NGS End Date to be benefit end date.
        'if no NGS End Date then use Category Effective date
        If IsDate(dlpDate(25).Text) Then
            xtmpdate = dlpDate(25).Text 'NGS End Date
        Else
            xtmpdate = dlpDate(34).Text 'Category Effective
        End If
        If clpPT.Text = "PT" Then
            If glbMsgCustomVal = 4 Then 'US NGS employee
                Call WFCUpdateBenefitEndDate(glbLEE_ID, xtmpdate, "ALL")
            End If
            If glbMsgCustomVal = 5 Then 'US NGS employee
                Call WFCUpdateBenefitEndDate(glbLEE_ID, xtmpdate, "ComPaidNoIE")
            End If
        End If
        If clpPT.Text = "FT" Then 'Ticket #25178 Franks 03/11/2014
            If glbMsgCustomVal = 8 Then  ' 7 Then 'US NGS employee
                'For US Benefit
                '"   Changed this employee from PT to FT. The NGS End Date and the Benefit End Dates were not removed.
                Call WFCUpdateBenefitEndDate(glbLEE_ID, xtmpdate, "RemoveEndDate")
            End If
        End If
        If Not (clpPT.Text = "PT" Or clpPT.Text = "FT") Then
            If glbMsgCustomVal = 4 Then 'US NGS employee
                Call WFCUpdateBenefitEndDate(glbLEE_ID, xtmpdate, "ALL")
            End If
        End If
    End If
End If

If glbtermopen Then
'   Hemu 07/02/2003 Begin - Ticket #4247, Update Employment Equity Data with Date of Termination
    If Len(dlpTermDate.Text) <> 0 Then
        'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
        'gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_DOT = " & Date_SQL(dlpTermDate.Text) & " WHERE EQ_EMPNBR = " & glbTERM_ID
        gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_DOT = " & Date_SQL(dlpTermDate.Text) & ", EQ_TYPE = 'T' WHERE EQ_EMPNBR = " & glbTERM_ID
    End If
'   Hemu 07/02/2003 End - Ticket #4247
    
    If glbWFC Then
        If IsDate(OTermDate) And IsDate(dlpTermDate) Then
            If Not (CVDate(OTermDate) = CVDate(dlpTermDate)) Then
                Call WFCPensionAlerts(lblEEID, Date, "Termination Date Change", dlpTermDate.Text, OTermDate, glbTERM_Seq)
            End If
        End If
    End If
End If
If Len(clpBGroup.Text) < 1 Then
    SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & lblEEID
    gdbAdoIhr001.Execute SQLQ
End If

If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Ticket #14752
    Call UptBenEndDate4StJohns
End If
    
    
If NewHireForms.count > 0 Then 'New Hire only
    Call EntReCalcPeriod("ED_EMPNBR=" & glbLEE_ID, "VAC")
    Call EntReCalcPeriod("ED_EMPNBR=" & glbLEE_ID, "SICK")
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    If NewHireForms.count > 0 Then 'New Hire only
        If gsEMAIL_ONNEWHIRE Then
            Call EmailSendingForSamuel
        End If
    End If
End If

If glbtermopen Then
    Call Employee_Master_Integration(lblEEID, , , glbTERM_Seq)
Else
    Call Employee_Master_Integration(glbLEE_ID)
    If glbMediPay Then Call Employee_Benefit_Integration(glbLEE_ID) 'Ticket #15201
End If

If glbGP Then 'Ticket #26654 Franks 02/11/2015
    If glbCompSerial = "S/N - 2453W" Then  'Town of Gander
        If SaveBGroup <> clpBGroup.Text Then
            Call Employee_GP_NewBenefitDeduction_Integration(glbLEE_ID)
        End If
    End If
End If

fglbNew = False

'    'George Jan 26,2006
'    gdbAdoIhr001_DOC.BeginTrans
'    gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_EMP set ??? where RE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
'    gdbAdoIhr001_DOC.CommitTrans
'    'George Jan 26,2006
'    'George Jan 26,2006
'    gdbAdoIhr001_DOC.BeginTrans
'    gdbAdoIhr001_DOC.Execute "Update HRDOC_EMP set ??? where RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR = " & glbLEE_ID
'    gdbAdoIhr001_DOC.CommitTrans
'    'George Jan 26,2006
       

Call SET_UP_MODE
'Call modSTUPD(True)

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Then
    If NewHireForms.count > 0 Then      'New Hire only
        Dim rsEmp As New ADODB.Recordset
        'Frank Ticket# 6647, Added ED_EMPNBR
        rsEmp.Open "SELECT ED_EMPNBR,ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmp.EOF Then
            If clpCode(2).Text = "NU" Then
                rsEmp("ED_SECTION") = "NU"
            ElseIf clpCode(2).Text = "NUE" Then
                rsEmp("ED_SECTION") = "EWS"
            Else
                rsEmp("ED_SECTION") = "U"
            End If
            rsEmp.Update
        End If
        rsEmp.Close
    End If
End If
'Hemu - End

'Added by Bryan 18/07/05 for Ticket #8963 Assemby of first Nations
If glbCompSerial = "S/N - 2376W" Then
    If Len(dlpDate(1).Text) > 0 Then
        SQLQ = "INSERT INTO HR_FOLLOW_UP(EF_COMPNO, EF_EMPNBR, EF_COMPLETED, EF_FREAS_TABL, EF_FREAS, EF_LDATE, EF_LTIME, EF_LUSER) "
        SQLQ = SQLQ + "VALUES (001, " & glbLEE_ID & ", 0, 'FURE', 'TE', " & Date_SQL(Now) & ", '" & Time$ & "', '" & glbUserID & "')"
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    If Len(dlpDate(2).Text) > 0 Then
        SQLQ = "INSERT INTO HR_FOLLOW_UP(EF_COMPNO, EF_EMPNBR, EF_COMPLETED, EF_FREAS_TABL, EF_FREAS, EF_LDATE, EF_LTIME, EF_LUSER) "
        SQLQ = SQLQ + "VALUES (001, " & glbLEE_ID & ", 0, 'FURE', 'PEN', " & Date_SQL(Now) & ", '" & Time$ & "', '" & glbUserID & "')"
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
End If
'Bryan - End

''Town of Aurora
''If glbCompSerial = "S/N - 2378W" Then
'If SavOrg <> clpCode(2).Text Then
'    Call Update_Overtime_Bank(clpCode(2).Text)
'End If
''End If

If NewHireForms.count > 0 Then  'New Hire only
    If Not glbtermopen Then
        Call UPDEML
        Call UPDOvertime_Overview
    End If
    
    If glbCompSerial = "S/N - 2387W" Then ' FOR Bird Packaging Limited Ticket #13701
        Call updFollowPension
    End If
    
    'Ticket #16395 Create WFC Pension Master record
    'New Hire
    If glbWFC Then
        If comEmpType.Text = "Y - Yes" Then
            toSOURCE = "IHR Status/Dates" 'Ticket #19954
            UpdPenAudDirect = True
            Call WFCPensionMaster(glbLEE_ID)
        End If
    End If
Else
    'Ticket #22847 - Add the call to procedure UPDOvertime_Overview to add Overtime Master record for employees whose
    'Employment Status, Category or Union changes. And doing the same from Demographics screen when Location, Region,
    'Admin By or Section changes.
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
    If Not glbtermopen Then
        If SavOrg <> clpCode(2).Text Or SavPT <> clpPT.Text Or SavEmp <> clpCode(1).Text Then   'Ticket #22847 - add PT and Status checking
            Call UPDOvertime_Overview   'Ticket #22847
            Call Update_Overtime_Bank(clpCode(2).Text)
        End If
    End If
    'End If

    'Ticket #16395 Create WFC Pension Master record
    'comEmpType change from "N" to "Y"
    If glbWFC Then
        If comEmpType.Text = "Y - Yes" Then
            If Not Left(OEmptype, 1) = "Y" Then
                If IsDate(dlpDate(8).Text) Then
                    xtmpdate = dlpDate(8).Text
                Else
                    xtmpdate = Date
                End If
                toSOURCE = "IHR Status/Dates" 'Ticket #19954
                UpdPenAudDirect = True
                Call WFCPensionMaster(glbLEE_ID, "Y", , , Year(Date)) 'xTmpDate
            End If
        End If
        If comEmpType.Text = "N - No" Then
            If Left(OEmptype, 1) = "Y" Then
                glbChgTermDate = ""
                frmMsgTerm.PenTermDate = "Y"
                frmMsgTerm.Show 1
                If IsDate(glbChgTermDate) Then
                    'Call WFCPensionMaster(glbLEE_ID, , glbChgTermDate)
                    toSOURCE = "IHR Status/Dates" 'Ticket #19954
                    Call WFCPensionMasUpt(glbLEE_ID, "PenExitDate", glbChgTermDate, , Year(glbChgTermDate))
                End If
            End If
        End If
    End If
End If


'Ticket #16395 Create WFC Pension Master record
If glbWFC Then
    If comEmpType.Text = "Y - Yes" Then
        'FT/PT/SE/TR/OT Change:
        If Len(Trim(SavPT)) > 0 And SavPT <> clpPT.Text Then
            'Pension Alert - Benficiary
            Call WFCPensionAlerts(glbLEE_ID, Date, "FT/PT Indicator Changed", clpPT.Text, SavPT)
        End If
        'Union Code Change
        If Len(Trim(SavOrg)) > 0 And SavOrg <> clpCode(2).Text Then
            'Pension Alert - Benficiary
            Call WFCPensionAlerts(glbLEE_ID, Date, "Union Change", clpCode(2).Text, SavOrg)
        End If
        'Original Hire Date Change
        If Len(Trim(SavDOH)) > 0 Then
            If Not (CVDate(SavDOH) = CVDate(dlpDate(7).Text)) Then
                toSOURCE = "IHR Status/Dates" 'Ticket #19954
                Call WFCPensionMasUpt(glbLEE_ID, "DOH_Change", dlpDate(7).Text, SavDOH, Year(Date))
            End If
        End If
        'Others -> STD
        If clpCode(1).Text = "STD" Then
            If SavEmp <> clpCode(1).Text Then
                If IsDate(dlpDate(15)) Then
                    Call Upt_PENSIONDATE2(glbLEE_ID, "UPDATE", dlpDate(15))
                End If
            End If
        End If
        'STD -> Others
        'If the Employment Status goes from STD to any status but LTD or DIS,
        'delete the Disability Date (ER_PENSIONDATE2)
        If SavEmp = "STD" Then
            If SavEmp <> clpCode(1).Text Then
                If Not (clpCode(1).Text = "LTD" Or clpCode(1).Text = "DIS") Then
                    Call Upt_PENSIONDATE2(glbLEE_ID, "DELETE")
                End If
            End If
        End If
        
        'Ticket #21597 Franks 05/01/2012
        '"   Complete = N, Event Date = 1st of the month following their 65 birthday minus 3 months, Event Type = "Retirement Notification"
        If NewHireForms.count > 0 Then
            xtmpdate = WFCPenEmpRetireDate(65, rsDATA("ED_DOB"))
            If IsDate(xtmpdate) Then
                xtmpdate = DateAdd("M", -3, xtmpdate)
                Call WFCPensionAlerts(glbLEE_ID, xtmpdate, "Retirement Notification")
            End If
        End If
    End If
    
    'Ticket #19266
    Call AUDIT_NGS_TRANS
End If
    
'Ticket #18790 - Create EEO record
If glbEmpCountry = "U.S.A." Then
    If NewHireForms.count > 0 Then  'For New Hires Only
        Call uptEEO_Fields(glbLEE_ID, "New")
    Else
        'if ft/pt or DOH changed
        'Ticket #20360 Franks 05/20/2011
        If Not (SavPT = clpPT.Text) Then
            Call uptEEO_Fields(glbLEE_ID, "Update")
        End If
        If Len(SavDOH) > 0 Then 'Ticket #20360 Franks 05/20/2011
            If Not (CVDate(SavDOH) = CVDate(dlpDate(7).Text)) Then
                Call uptEEO_Fields(glbLEE_ID, "Update")
            End If
        End If
    End If
End If

Me.clpCode(1).SetFocus
Call NextForm

If glbCompSerial = "S/N - 2217W" Then ' FOR CITY OF PICKERING
    If Not updFollow("U") Then Exit Sub
End If
If glbCompSerial = "S/N - 2352W" Then ' FOR Tobias House
    Call updFollowSin
End If
If glbWFC Then
    Call updFollowUserDefinedDate
End If

'Release 8.0 - Ticket #22682: Create a Follow Up record when the Employment Status's To Date is entered
If OTDate <> dlpDate(16).Text Then
    Call updFollowStatusToDate
End If

Call EERetrieve

If glbWFC And glbMsgCustomVal = 2 Then 'Ticket #23903 Franks 06/19/2013
    'Union Moving - call transfer out
    Call OpenTranOutForm
    Unload Me
End If

''If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
''    Call FamilyDayEmpSync(glbLEE_ID)
''End If
'Ticket #24729 01/28/2014 Franks - users need to change this form for both ids

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Step" & Trim(Str(Y%)))

Resume Next
Unload Me

End Sub

Private Sub EmailSendingForSamuel()
Dim rsLEmp As New ADODB.Recordset
Dim MailBody As String
Dim SQLQ As String
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String
Dim xEmpName As String

On Error GoTo Email_Err
    
    MailBody = GetEmailBodyForSamuel(glbLEE_ID)
    MailBody = MailBody & "was hired on " & dlpDate(7).Text & "."
    
    'Ticket #24685 - Adding French to the New Hire email - Begin
    xEmpName = GetEmpData(glbLEE_ID, "ED_FNAME", "") & " " & GetEmpData(glbLEE_ID, "ED_SURNAME", "")
    MailBody = MailBody & vbCrLf
    MailBody = MailBody & "L'employé(e) #" & glbLEE_ID & " - " & xEmpName
    MailBody = MailBody & ", sur la paie " & GetEmpData(glbLEE_ID, "ED_ADMINBY", "")
    MailBody = MailBody & ", de la Succursale " & GetTABLDesc("EDSE", GetEmpData(glbLEE_ID, "ED_SECTION", ""))
    MailBody = MailBody & " a été embauché(e) le " & dlpDate(7).Text & "."
    'Ticket #24685 - Adding French to the New Hire email - End
    
    xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", glbLEE_ID)
    If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
        xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
    End If
    If Len(xToEmail) > 0 Then
        frmSendEmail.txtTo.Text = xToEmail
        'frmSendEmail.txtSubject.Text = "info:HR Employee New Hire Notice"
        'Ticket #18578
        xBranch = GetTABLDesc("EDSE", GetEmpData(glbLEE_ID, "ED_SECTION", ""))
        If Len(xBranch) > 0 Then
            xBranch = xBranch & " - "
        End If
        
        'Ticket #24685 - Adding French to the New Hire email - Begin
        'xEmailSubject = "info:HR Employee New Hire Notice - " & xBranch & frmEESTATS.lblEEName
        xEmailSubject = "info:HR Employee New Hire Notice - " & xBranch & frmEESTATS.lblEEName
        xEmailSubject = xEmailSubject & " \ "
        xEmailSubject = xEmailSubject & "info:HR - Avis d'Embauche Nouvel(le) Employé(e) - " & xBranch & frmEESTATS.lblEEName
        xEmailSubject = xEmailSubject & "Courriel - Avis info:HR"
        'Ticket #24685 - Adding French to the New Hire email - End
        
        frmSendEmail.txtSubject.Text = xEmailSubject
        frmSendEmail.txtBody.Text = MailBody
        'frmSendEmail.Show 1
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = "Sending email..."
        frmSendEmail.Tag = ""
        frmSendEmail.cmdSend_Click
        Do
            DoEvents
        Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
        ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
        If frmSendEmail.Tag = "DONE" Then
            Unload frmSendEmail
        Else
            Unload frmSendEmail
        End If
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
    End If

Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    'Resume Next
    Exit Sub
End Sub

Private Sub Update_Overtime_Bank(xUnion)
Dim rsOvtEmp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String

'Recalculate the Overtime Bank
SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
rsOvtEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsOvtEmp.EOF Then
    'Ticket #22847- Employee Overtime Bank record taken care from UPDOvertime_Overview function
    'If xUnion = "" Then
        'Delete the record from Overtime Bank
    '    gdbAdoIhr001.Execute "DELETE HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
    'Else
        Call ReCalcOvt("OT_EMPNBR = " & glbLEE_ID)
    'End If
Else
    If xUnion <> "" Then
'Let the user run the "Update All Employees" from Overtime Bank Master screen to add
'        SQLQ = "SELECT ED_EMPNBR, ED_ORG FROM HREMP "
'        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
'        SQLQ = SQLQ & " AND ED_ORG IN (SELECT OM_ORG FROM HR_OVERTIME_MASTER)"
'        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'        If Not rsEMP.EOF Then
'            'Add record Overtime Bank records
'            rsOvtEmp.AddNew
'            rsOvtEmp("OT_COMPNO") = "001"
'            rsOvtEmp("OT_EMPNBR") = glbLEE_ID
'            rsOvtEmp("OT_PBANK") = 0
'            rsOvtEmp("OT_BANK") = Get_OvertimeBank(glbLEE_ID) * Overtime_Multiplier(xUnion, clpPT.Text, clpCode(1).Text)
'            rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(glbLEE_ID)
'            rsOvtEmp("OT_EFDATE") = Format("1/1/" & Year(Now()), "mm/dd/yyyy")
'            rsOvtEmp("OT_ETDATE") = Format("12/31/" & Year(Now()), "mm/dd/yyyy")
'            rsOvtEmp("OT_LDATE") = Date
'            rsOvtEmp("OT_LTIME") = Time$
'            rsOvtEmp("OT_LUSER") = glbUserID
'            rsOvtEmp.Update
'        End If
'        rsEMP.Close
    End If
End If
rsOvtEmp.Close


End Sub

Private Sub PayAudit_TermNewhire(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTPayID
        If Len(glbChgNewEmpnbr) = 0 Then Exit Sub
        rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsTC.EOF Then Exit Sub
        If IsNull(rsTC("ED_PAYROLL_ID")) Then Exit Sub
        xTPayID = rsTC("ED_PAYROLL_ID")
        
        'Termination Data
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_SURNAME") = rsTC("ED_SURNAME") '
        rsTA("AU_FNAME") = rsTC("ED_FNAME")
        rsTA("AU_DOT") = glbChgTermDate
        rsTA("AU_TREAS") = glbChgTermReason
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_PAYROLL_ID") = xTPayID
        rsTA("AU_DIVUPL") = rsTC("ED_DIV")
        rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "T"
        rsTA.Update
        'rsTC.Close
 
        'New Hire Data (Payroll ID may change)
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1": rsTA("AU_LANG2_TABL") = "EDL1"
        rsTA("AU_DIV") = rsTC("ED_DIV")
        rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
        rsTA("AU_TITLE") = rsTC("ED_TITLE")
        rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
        rsTA("AU_FNAME") = rsTC("ED_FNAME")
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_PAYROLL_ID") = glbChgNewEmpnbr
        rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
        rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
        rsTA("AU_CITY") = rsTC("ED_CITY")
        rsTA("AU_PROV") = rsTC("ED_PROV")
        rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
        rsTA("AU_PCODE") = rsTC("ED_PCODE")
        rsTA("AU_PHONE") = rsTC("ED_PHONE")
        rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
        rsTA("AU_DIVUPL") = rsTC("ED_DIV")
        rsTA("AU_SEX") = rsTC("ED_SEX")
        rsTA("AU_SMOKER") = IIf(rsTC("ED_SMOKER"), "Yes", "No")
        rsTA("AU_DOB") = rsTC("ED_DOB")
        rsTA("AU_SIN") = rsTC("ED_SIN")
        rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
        rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
        rsTA("AU_NEWEMP") = "Y"
        rsTA("AU_PTUPL") = rsTC("ED_PT")
        rsTA("AU_LOC") = rsTC("ED_LOC")
        rsTA("AU_TD1") = rsTC("ED_TD1")
        rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
        rsTA("AU_PROVFORM") = rsTC("ED_PROVFORM")
        rsTA("AU_PROVAMT") = rsTC("ED_PROVAMT")
        rsTA("AU_OLDTD1") = 0
        rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
        rsTA("AU_REGION") = rsTC("ED_REGION")
        rsTA("AU_SECTION") = rsTC("ED_SECTION")
        rsTA("AU_HOMEOPRTNBR") = rsTC("ED_HOMEOPRTNBR")
        rsTA("AU_HOMELINE") = rsTC("ED_HOMELINE")
        rsTA("AU_HOMESHIFT") = rsTC("ED_HOMESHIFT")
        rsTA("AU_HOMEWRKCNT") = rsTC("ED_HOMEWRKCNT")
        rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
        rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
        rsTA("AU_SSN") = rsTC("ED_SSN")
     
        rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
        rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
        rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
        rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
        rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
        rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
        rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
        rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")
        rsTA("AU_BADGEID") = rsTC("ED_BADGEID")
        rsTA("AU_MIDNAME") = rsTC("ED_MIDNAME")
        rsTA("AU_ALIAS") = rsTC("ED_ALIAS")
        'Employee Status
        rsTA("AU_EMP") = clpCode(1) 'rsTC("ED_EMP")
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
        rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
        rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
        rsTA("AU_BANK") = rsTC("ED_BANK")
        rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
        rsTA("AU_TRANSITABA") = rsTC("ED_TRANSITABA")
        rsTA("AU_TRANSITABA2") = rsTC("ED_TRANSITABA2")
        rsTA("AU_TRANSITABA3") = rsTC("ED_TRANSITABA3")
        rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
        rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
        'BANK 2
        rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
        rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
        rsTA("AU_BANK2") = rsTC("ED_BANK2")
        rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
        rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
        'BANK3
        rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
        rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
        rsTA("AU_BANK3") = rsTC("ED_BANK3")
        rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
        rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
        rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
        
        rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
        rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
        rsTA("AU_TD3") = rsTC("ED_TD3")
        rsTA("AU_TD1") = rsTC("ED_TD1")
        rsTA("AU_DDI") = rsTC("ED_DDI")
        rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
        rsTA("AU_FedTax") = rsTC("ED_FedTax")
        rsTA("AU_ExtAmt") = rsTC("ED_ExtAmt")
        rsTA("AU_ProvForm") = rsTC("ED_ProvForm")
        rsTA("AU_ProvAmt") = rsTC("ED_ProvAmt")
        rsTA("AU_ExtraTax") = rsTC("ED_ExtraTax")
        rsTA("AU_ExtraTaxPC") = rsTC("ED_ExtraTaxPC")
    
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
        rsTA.Update
        rsTC.Close
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
            rsTA("AU_SEDATE") = rsTC("SH_NEXTDAT")
            rsTA("AU_PAYP") = rsTC("SH_PAYP")
        End If
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
        rsTA.Update
        rsTC.Close
        '------Job and Salary Information END
        
        '------Other Earnings Begin
        rsTC.Open "SELECT * FROM HREARN WHERE EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
        Do While Not rsTC.EOF
            rsTA.AddNew
            rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
            rsTA("AU_NEWEMP") = "N"
            rsTA("AU_EARN") = rsTC("EARN_TYPE")
            rsTA("AU_ADOLLAR") = rsTC("ACT_DOLLAR")
            rsTA("AU_COEFLAG") = IIf(rsTC("COST_OF_EMPLOYMENT"), "Y", "N")
            rsTA("AU_COMPNO") = "001"
            rsTA("AU_EMPNBR") = glbLEE_ID
            rsTA("AU_LDATE") = Date
            rsTA("AU_LUSER") = glbUserID
            rsTA("AU_LTIME") = Time$
            rsTA("AU_UPLOAD") = "N"
            rsTA("AU_TYPE") = "A"
            rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
            rsTA.Update
            rsTC.MoveNext
        Loop
        rsTC.Close
        '------Other Earnings End
        
        'Change the ED_PAYROLL_ID in HREMP
        rsTC.Open "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsTC.EOF Then
            rsTC("ED_PAYROLL_ID") = glbChgNewEmpnbr
            rsTC.Update
        End If
        rsTC.Close
        
End Sub
Private Sub Samuel_Audit(rsTA As ADODB.Recordset, xField, xVal, xLDate) 'Ticket #20600 Franks 09/02/2011
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTilPayID
    If Not IsDate(xLDate) Then Exit Sub
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If rsTC.EOF Then Exit Sub
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1": rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
    rsTA(xField) = xVal
    rsTA("AU_LDATE") = Format(xLDate, "SHORT DATE")
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA.Update
        
End Sub
Private Sub WFC_GREN_Audit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTilPayID
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If rsTC.EOF Then Exit Sub
    If IsNull(rsTC("ED_PAYROLL_ID")) Then Exit Sub
    xTilPayID = rsTC("ED_PAYROLL_ID")

    If Len(glbChgNewEmpnbr) = 0 Then 'Termination Data
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_SURNAME") = rsTC("ED_SURNAME") '
        rsTA("AU_FNAME") = rsTC("ED_FNAME")
        rsTA("AU_DOT") = glbChgTermDate
        rsTA("AU_TREAS") = glbChgTermReason
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_PAYROLL_ID") = xTilPayID
        rsTA("AU_DIVUPL") = rsTC("ED_DIV")
        rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "T"
        rsTA.Update
        rsTC.Close
    Else
        'New Hire Data (Payroll ID may change)
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1": rsTA("AU_LANG2_TABL") = "EDL1"
        rsTA("AU_DIV") = rsTC("ED_DIV")
        rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
        rsTA("AU_TITLE") = rsTC("ED_TITLE")
        rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
        rsTA("AU_FNAME") = rsTC("ED_FNAME")
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_PAYROLL_ID") = glbChgNewEmpnbr
        rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
        rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
        rsTA("AU_CITY") = rsTC("ED_CITY")
        rsTA("AU_PROV") = rsTC("ED_PROV")
        rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
        rsTA("AU_PCODE") = rsTC("ED_PCODE")
        rsTA("AU_PHONE") = rsTC("ED_PHONE")
        rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
        rsTA("AU_DIVUPL") = rsTC("ED_DIV")
        rsTA("AU_SEX") = rsTC("ED_SEX")
        rsTA("AU_SMOKER") = IIf(rsTC("ED_SMOKER"), "Yes", "No")
        rsTA("AU_DOB") = rsTC("ED_DOB")
        rsTA("AU_SIN") = rsTC("ED_SIN")
        rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
        rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
        rsTA("AU_NEWEMP") = "Y"
        rsTA("AU_PTUPL") = rsTC("ED_PT")
        rsTA("AU_LOC") = rsTC("ED_LOC")
        rsTA("AU_TD1") = rsTC("ED_TD1")
        rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
        rsTA("AU_PROVFORM") = rsTC("ED_PROVFORM")
        rsTA("AU_PROVAMT") = rsTC("ED_PROVAMT")
        rsTA("AU_OLDTD1") = 0
        rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
        rsTA("AU_REGION") = rsTC("ED_REGION")
        rsTA("AU_SECTION") = rsTC("ED_SECTION")
        rsTA("AU_HOMEOPRTNBR") = rsTC("ED_HOMEOPRTNBR")
        rsTA("AU_HOMELINE") = rsTC("ED_HOMELINE")
        rsTA("AU_HOMESHIFT") = rsTC("ED_HOMESHIFT")
        rsTA("AU_HOMEWRKCNT") = rsTC("ED_HOMEWRKCNT")
        rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
        rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
        rsTA("AU_SSN") = rsTC("ED_SSN")
     
        rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
        rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
        rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
        rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
        rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
        rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
        rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
        rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")
        rsTA("AU_BADGEID") = rsTC("ED_BADGEID")
        rsTA("AU_MIDNAME") = rsTC("ED_MIDNAME")
        rsTA("AU_ALIAS") = rsTC("ED_ALIAS")
        'Employee Status
        rsTA("AU_EMP") = clpCode(1) 'rsTC("ED_EMP")
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
        rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
        rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
        rsTA("AU_BANK") = rsTC("ED_BANK")
        rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
        rsTA("AU_TRANSITABA") = rsTC("ED_TRANSITABA")
        rsTA("AU_TRANSITABA2") = rsTC("ED_TRANSITABA2")
        rsTA("AU_TRANSITABA3") = rsTC("ED_TRANSITABA3")
        rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
        rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
        'BANK 2
        rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
        rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
        rsTA("AU_BANK2") = rsTC("ED_BANK2")
        rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
        rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
        'BANK3
        rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
        rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
        rsTA("AU_BANK3") = rsTC("ED_BANK3")
        rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
        rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
        rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
        
        rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
        rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
        rsTA("AU_TD3") = rsTC("ED_TD3")
        rsTA("AU_TD1") = rsTC("ED_TD1")
        rsTA("AU_DDI") = rsTC("ED_DDI")
        rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
        rsTA("AU_FedTax") = rsTC("ED_FedTax")
        rsTA("AU_ExtAmt") = rsTC("ED_ExtAmt")
        rsTA("AU_ProvForm") = rsTC("ED_ProvForm")
        rsTA("AU_ProvAmt") = rsTC("ED_ProvAmt")
        rsTA("AU_ExtraTax") = rsTC("ED_ExtraTax")
        rsTA("AU_ExtraTaxPC") = rsTC("ED_ExtraTaxPC")
    
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
        rsTA.Update
        rsTC.Close
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
            rsTA("AU_SEDATE") = rsTC("SH_NEXTDAT")
            rsTA("AU_PAYP") = rsTC("SH_PAYP")
        End If
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
        rsTA.Update
        rsTC.Close
        '------Job and Salary Information END
        
        '------Other Earnings Begin
        rsTC.Open "SELECT * FROM HREARN WHERE EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
        Do While Not rsTC.EOF
            rsTA.AddNew
            rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
            rsTA("AU_NEWEMP") = "N"
            rsTA("AU_EARN") = rsTC("EARN_TYPE")
            rsTA("AU_ADOLLAR") = rsTC("ACT_DOLLAR")
            rsTA("AU_COEFLAG") = IIf(rsTC("COST_OF_EMPLOYMENT"), "Y", "N")
            rsTA("AU_COMPNO") = "001"
            rsTA("AU_EMPNBR") = glbLEE_ID
            rsTA("AU_LDATE") = Date
            rsTA("AU_LUSER") = glbUserID
            rsTA("AU_LTIME") = Time$
            rsTA("AU_UPLOAD") = "N"
            rsTA("AU_TYPE") = "A"
            rsTA("AU_Payroll_ID") = glbChgNewEmpnbr
            rsTA.Update
            rsTC.MoveNext
        Loop
        rsTC.Close
        '------Other Earnings End
    End If


End Sub
Private Sub TilburyPayrollIDAudit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTilPayID
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If rsTC.EOF Then Exit Sub
    If IsNull(rsTC("ED_PAYROLL_ID")) Then Exit Sub
    xTilPayID = rsTC("ED_PAYROLL_ID")
    'Termination Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_SURNAME") = rsTC("ED_SURNAME") '
    rsTA("AU_FNAME") = rsTC("ED_FNAME")
    rsTA("AU_DOT") = glbChgTermDate
    rsTA("AU_TREAS") = glbChgTermReason
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_PAYROLL_ID") = "8" & xTilPayID
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
    rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    rsTA.Update
    
    'New Hire Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1": rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_DIV") = rsTC("ED_DIV")
    rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
    rsTA("AU_TITLE") = rsTC("ED_TITLE")
    rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
    rsTA("AU_FNAME") = rsTC("ED_FNAME")
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_PAYROLL_ID") = xTilPayID
    rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
    rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
    rsTA("AU_CITY") = rsTC("ED_CITY")
    rsTA("AU_PROV") = rsTC("ED_PROV")
    rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
    rsTA("AU_PCODE") = rsTC("ED_PCODE")
    rsTA("AU_PHONE") = rsTC("ED_PHONE")
    rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
    rsTA("AU_SEX") = rsTC("ED_SEX")
    rsTA("AU_SMOKER") = IIf(rsTC("ED_SMOKER"), "Yes", "No")
    rsTA("AU_DOB") = rsTC("ED_DOB")
    rsTA("AU_SIN") = rsTC("ED_SIN")
    rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
    rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
    rsTA("AU_NEWEMP") = "Y"
    rsTA("AU_PTUPL") = rsTC("ED_PT")
    rsTA("AU_LOC") = rsTC("ED_LOC")
    rsTA("AU_TD1") = rsTC("ED_TD1")
    rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
    rsTA("AU_PROVFORM") = rsTC("ED_PROVFORM")
    rsTA("AU_PROVAMT") = rsTC("ED_PROVAMT")
    rsTA("AU_OLDTD1") = 0
    rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
    rsTA("AU_REGION") = rsTC("ED_REGION")
    rsTA("AU_SECTION") = rsTC("ED_SECTION")
    rsTA("AU_HOMEOPRTNBR") = rsTC("ED_HOMEOPRTNBR")
    rsTA("AU_HOMELINE") = rsTC("ED_HOMELINE")
    rsTA("AU_HOMESHIFT") = rsTC("ED_HOMESHIFT")
    rsTA("AU_HOMEWRKCNT") = rsTC("ED_HOMEWRKCNT")
    rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
    rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
    rsTA("AU_SSN") = rsTC("ED_SSN")
 
    rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
    rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
    rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
    rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
    rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
    rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
    rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
    rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")
    rsTA("AU_BADGEID") = rsTC("ED_BADGEID")
    rsTA("AU_MIDNAME") = rsTC("ED_MIDNAME")
    rsTA("AU_ALIAS") = rsTC("ED_ALIAS")
    'Employee Status
    rsTA("AU_EMP") = clpCode(1) 'rsTC("ED_EMP")
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
    rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
    rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
    rsTA("AU_BANK") = rsTC("ED_BANK")
    rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
    rsTA("AU_TRANSITABA") = rsTC("ED_TRANSITABA")
    rsTA("AU_TRANSITABA2") = rsTC("ED_TRANSITABA2")
    rsTA("AU_TRANSITABA3") = rsTC("ED_TRANSITABA3")
    rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
    rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
    'BANK 2
    rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
    rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
    rsTA("AU_BANK2") = rsTC("ED_BANK2")
    rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
    rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
    'BANK3
    rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
    rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
    rsTA("AU_BANK3") = rsTC("ED_BANK3")
    rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
    rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
    rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
    
    rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
    rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
    rsTA("AU_TD3") = rsTC("ED_TD3")
    rsTA("AU_TD1") = rsTC("ED_TD1")
    rsTA("AU_DDI") = rsTC("ED_DDI")
    rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
    rsTA("AU_FedTax") = rsTC("ED_FedTax")
    rsTA("AU_ExtAmt") = rsTC("ED_ExtAmt")
    rsTA("AU_ProvForm") = rsTC("ED_ProvForm")
    rsTA("AU_ProvAmt") = rsTC("ED_ProvAmt")
    rsTA("AU_ExtraTax") = rsTC("ED_ExtraTax")
    rsTA("AU_ExtraTaxPC") = rsTC("ED_ExtraTaxPC")

    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA("AU_Payroll_ID") = xTilPayID
    rsTA.Update
    rsTC.Close
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
        rsTA("AU_SEDATE") = rsTC("SH_NEXTDAT")
        rsTA("AU_PAYP") = rsTC("SH_PAYP")
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA("AU_Payroll_ID") = xTilPayID
    rsTA.Update
    rsTC.Close
    '------Job and Salary Information END
    
    '------Other Earnings Begin
    rsTC.Open "SELECT * FROM HREARN WHERE EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    Do While Not rsTC.EOF
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_EARN") = rsTC("EARN_TYPE")
        rsTA("AU_ADOLLAR") = rsTC("ACT_DOLLAR")
        rsTA("AU_COEFLAG") = IIf(rsTC("COST_OF_EMPLOYMENT"), "Y", "N")
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = xTilPayID
        rsTA.Update
        rsTC.MoveNext
    Loop
    rsTC.Close
    '------Other Earnings End

End Sub
'Private Function ChkDupBenCode()
'Dim rsTmp As New ADODB.Recordset
'Dim xFlag As Boolean
'Dim SQLQ
'    xFlag = False
'
'
'End Function

Private Sub UptBenEndDate4StJohns()
Dim SQLQ As String
    If clpCode(1).Text = "TERM" Then
        If IsDate(dlpDate(1).Text) Then 'Last Day
            SQLQ = "UPDATE HRBENFT SET BF_CEASEDATE = " & Date_SQL(dlpDate(1).Text) & " WHERE BF_EMPNBR = " & glbLEE_ID
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
End Sub

Private Sub UpdateGPMainBenDed()
Dim x%, Y%, SQLQ, Msg
Dim wInComdOld As Boolean
Dim wInComdNew As Boolean

    'Ticket #17336 - County of Lanark, GP Benefit/Deduction Pay Code Integation
    'The the following GP Integration function must be turn on
    'Employee Pay Code - Salary Based Pay Codes
    'Employee Pay Code - Benefit/Deduction Pay Codes
    If Not glbGP Then
        Exit Sub
    End If
    If NewHireForms.count > 0 Then
        'No position and salary, can't create the Pay Code records
        Exit Sub
    End If
    'If isTransferGP("Great Plains", "Emp_PayCode_Benefit_To_GP") Or isTransferGP("Great Plains", "Emp_PayCode_Salary_To_GP") Then
    If isTransferGP("Great Plains", "Emp_PayCode_Salary_To_GP") Then
        If SavOrg <> clpCode(2).Text Then
            'check if there are Pay codes associated with this union code
            wInComdOld = GPBDPayCode(SavOrg)
            wInComdNew = GPBDPayCode(clpCode(2).Text)
            If wInComdNew Or wInComdOld Then
                  If wInComdNew Then 'Add, update, delete
                    'Msg = "Do you want add/update the Employee's Pay Codes with the Benefit/Deduction " & Chr(10)
                    Msg = "Do you want add/update the Employee's Pay Codes with the Income Codes " & Chr(10)
                    Msg = Msg & "defined in the Income Code Matrix under menu item Great Plains? "
                    If MsgBox(Msg, 36, "info:HR") = 6 Then
                        Call UpdateGPBenefitDeduction(glbLEE_ID, clpCode(2).Text, SavOrg)
                        DoEvents
                        frmGPPayCodeList.Show 1
                    End If
                Else
                    If wInComdOld Then 'delete the old income codes only
                        Msg = "Do you want delete the Employee's Pay Codes with the Benefit/Deduction " & Chr(10)
                        Msg = Msg & "Codes '" & SavOrg & "' defined in the Income Code Matrix under menu item Great Plains? "
                        If MsgBox(Msg, 36, "info:HR") = 6 Then
                            Call UpdateGPBenefitDeduction(glbLEE_ID, clpCode(2).Text, SavOrg)
                            'Call Employee_GP_BenefitDeduction_Integration(glbLEE_ID, glbUserID, True)
                            DoEvents
                            frmGPPayCodeList.Show 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub UpdateBenefitGroup()
Dim rsBGMST As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String

Dim BelongOldGroup As Boolean
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
gdbAdoIhr001W.CommitTrans

'Len(NewBGroup) = 0, it means deleting the Benefit Group
If Len(NewBGroup) > 0 Then
    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "' "
    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_COMPNO") = "001"
        rsBGTMP("BM_BENEFIT_GROUP") = NewBGroup
        rsBGTMP("BM_BCODE") = rsBGMST("BM_BCODE")
        If glbWFC Then  'Ticket #18654 06/11/2010 Frank
            If NewHireForms.count > 0 Then
                rsBGTMP("BM_EDATE") = dlpDate(7).Text
                If Not IsNull(rsBGMST("BM_WAITPERIOD")) Then
                    If rsBGMST("BM_WAITPERIOD") > 0 Then
                        rsBGTMP("BM_EDATE") = CountEDate(lblEEID.Caption, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), dlpDate(7).Text)
                    End If
                End If
            Else
                If glbEmpCountry = "CANADA" Then
                    rsBGTMP("BM_EDATE") = Date
                Else
                    'Ticket #24937 Franks 02/07/2014 - for non CANADA employees
                    rsBGTMP("BM_EDATE") = dlpDate(7).Text
                    If Not IsNull(rsBGMST("BM_WAITPERIOD")) Then
                        If rsBGMST("BM_WAITPERIOD") > 0 Then
                            rsBGTMP("BM_EDATE") = CountEDate(lblEEID.Caption, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), dlpDate(7).Text)
                        End If
                    End If
                End If
            End If
            If rsBGMST("BM_BCODE") = "LIF1" Or rsBGMST("BM_BCODE") = "LIF3" Then
                rsBGTMP("BM_CHECK") = 0
            Else
                rsBGTMP("BM_CHECK") = 1
            End If
        Else
            If IsDate(rsBGMST("BM_EDATE")) Then
                rsBGTMP("BM_EDATE") = rsBGMST("BM_EDATE")
            Else
                'Ticket #24203 - Family Day Care Services
                'Ticket #21504 - Kerry's Place
                If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2436W" Then
                    'Passing User Defined Date instead of Hire Date
                    rsBGTMP("BM_EDATE") = CountEDate(lblEEID.Caption, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), dlpDate(3).Text)
                Else
                    'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
                    'Passing Hire Date incase if new hire
                    If glbCompSerial = "S/N - 2420W" And rsBGMST("BM_BCODE") = "PEN" Then
                        rsBGTMP("BM_EDATE") = CountEDate(lblEEID.Caption, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), dlpDate(7).Text, , rsBGMST("BM_BCODE"))
                    Else
                        'Passing Hire Date incase if new hire
                        rsBGTMP("BM_EDATE") = CountEDate(lblEEID.Caption, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), dlpDate(7).Text)
                    End If
                End If
            End If
            rsBGTMP("BM_CHECK") = 1
        End If
        rsBGTMP("BM_COVER") = rsBGMST("BM_COVER")
        rsBGTMP("BM_AMT") = rsBGMST("BM_AMT")
        rsBGTMP("BM_PPAMT") = rsBGMST("BM_PPAMT")
        rsBGTMP("BM_UNITCOST") = rsBGMST("BM_UNITCOST")
        rsBGTMP("BM_PCE") = rsBGMST("BM_PCE")
        rsBGTMP("BM_PCC") = rsBGMST("BM_PCC")
        rsBGTMP("BM_ECOST") = rsBGMST("BM_ECOST")
        rsBGTMP("BM_CCOST") = rsBGMST("BM_CCOST")
        rsBGTMP("BM_TCOST") = rsBGMST("BM_TCOST")
        rsBGTMP("BM_MAXDOL") = rsBGMST("BM_MAXDOL")
        rsBGTMP("BM_PREMIUM") = rsBGMST("BM_PREMIUM")
        rsBGTMP("BM_PER") = rsBGMST("BM_PER")
        rsBGTMP("BM_MTHCCOST") = rsBGMST("BM_MTHCCOST")
        rsBGTMP("BM_MTHECOST") = rsBGMST("BM_MTHECOST")
        rsBGTMP("BM_TAXBEN") = rsBGMST("BM_TAXBEN")
        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BM_SALARYDEPENDANT")
        rsBGTMP("BM_MINIMUM") = rsBGMST("BM_MINIMUM")
        rsBGTMP("BM_FACTOR") = rsBGMST("BM_FACTOR")
        rsBGTMP("BM_ROUND") = rsBGMST("BM_ROUND")
        rsBGTMP("BM_MAXIMUM") = rsBGMST("BM_MAXIMUM")
        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BM_NEXTNEAREST")
        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BM_TAXAMOUNT")
        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
        
        rsBGTMP("BM_DWM") = rsBGMST("BM_DWM")
        rsBGTMP("BM_PERORDOLL") = rsBGMST("BM_PERORDOLL")
        
        rsBGTMP("BM_POLICY") = rsBGMST("BM_POLICY")
        
        'Ticket #20931 - Rate Level
        rsBGTMP("BM_RATELEVEL") = rsBGMST("BM_RATELEVEL")
        
        rsBGTMP("BM_COMMENTS") = rsBGMST("BM_COMMENTS")
        rsBGTMP("BM_PTAX") = rsBGMST("BM_PTAX")
        'rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_ACTION") = "Add"
        rsBGTMP("BM_WRKEMP") = glbUserID
        
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BM_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
    If Not glbSQL And Not glbOracle Then Call Pause(1)
    
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & lblEEID
    
    rsBGEE.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    gdbAdoIhr001W.BeginTrans
    Do Until rsBGEE.EOF
        SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID
        SQLQ = SQLQ & "' AND  BM_BCODE='" & rsBGEE("BF_BCODE") & "'"
        SQLQ = SQLQ & " AND BM_ACTION='Add' " 'Frank 11/04/2003 for Duplicate record entere and delete
        'Ticket #24178 Franks 08/06/2013 - begin
        If IsNull(rsBGEE("BF_COVER")) Then
            SQLQ = SQLQ & " AND (BM_COVER IS NULL OR BM_COVER='')"
        Else
            SQLQ = SQLQ & " AND BM_COVER='" & rsBGEE("BF_COVER") & "'"
        End If
        'Ticket #24178 Franks 08/06/2013 - end
        rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
        If rsBGTMP.EOF Then
        'If rsBGTMP.EOF Or rsBGEE("BF_GROUP") = SaveBGroup Then
            BelongOldGroup = False
            If rsBGEE("BF_GROUP") = SaveBGroup Then
                BelongOldGroup = True
            Else
                SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & SaveBGroup & "'"
                SQLQ = SQLQ & " AND  BM_BCODE='" & rsBGEE("BF_BCODE")
                If IsNull(rsBGEE("BF_COVER")) Then
                    SQLQ = SQLQ & "' AND (BM_COVER IS NULL OR BM_COVER='')"
                Else
                    SQLQ = SQLQ & "' AND BM_COVER='" & rsBGEE("BF_COVER") & "'"
                End If
                rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsBGMST.EOF Then
                    BelongOldGroup = True
                End If
                rsBGMST.Close
            End If
            If BelongOldGroup Then
                rsBGTMP.AddNew
                rsBGTMP("BM_BCODE") = rsBGEE("BF_BCODE")
                rsBGTMP("BM_COVER") = rsBGEE("BF_COVER")
                rsBGTMP("BM_EDATE") = rsBGEE("BF_EDATE")
                rsBGTMP("BM_BENEFIT_GROUP") = SaveBGroup
                rsBGTMP("BM_CHECK") = 1
                If glbWFC Then 'Ticket #18810
                    'Cannot delete the Benefits since Manulife needs the EndDate
                    rsBGTMP("BM_ACTION") = "EndDate"
                    rsBGTMP("BM_BENEFIT_GROUP") = Null 'delete old Benefit Group
                Else
                    rsBGTMP("BM_ACTION") = "Delete"
                End If
                rsBGTMP("BM_WRKEMP") = glbUserID
                rsBGTMP("BM_WRKID") = rsBGEE("BF_BENE_ID")
                SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGEE("BF_BCODE") & "' "
                rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsTABL.EOF Then
                    rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
                End If
                rsTABL.Close
                rsBGTMP("BM_CHECK") = 1
                rsBGTMP.Update
            End If
            
        Else
            rsBGTMP("BM_WRKID") = rsBGEE("BF_BENE_ID")
            rsBGTMP("BM_CHECK") = 1
            rsBGTMP("BM_ACTION") = "Update"
            rsBGTMP.Update
        End If
        rsBGTMP.Close
        rsBGEE.MoveNext
    Loop
    gdbAdoIhr001W.CommitTrans
Else 'Deleting the Benefit Group
    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & lblEEID & " "
    SQLQ = SQLQ & "AND BF_GROUP ='" & SaveBGroup & "' "
    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
        rsBGTMP("BM_BENEFIT_GROUP") = SaveBGroup
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_ACTION") = "Delete"
        rsBGTMP("BM_WRKEMP") = glbUserID
        rsBGTMP("BM_WRKID") = rsBGMST("BF_BENE_ID")
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
End If

End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Status/Dates Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Status/Dates Information Report"
Me.vbxCrystal.Reset
Me.vbxCrystal.Destination = crptToPrinter
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Call setRptLabel(Me, 1)
If Not glbtermopen Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    xReport = glbIHRREPORTS & "rgstatus.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 7
            If x% = 2 Or x% = 6 Then
                Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
            Else
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next
    End If
    xReport = glbIHRREPORTS & "rgstatu2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Public Sub cmdView_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Status/Dates Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Status/Dates Information Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.Destination = crptToWindow
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Call setRptLabel(Me, 1)
If Not glbtermopen Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    xReport = glbIHRREPORTS & "rgstatus.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 7
            If x% = 2 Or x% = 6 Then
                Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
            Else
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next
    End If
    xReport = glbIHRREPORTS & "rgstatu2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Private Sub clpBGroup_LostFocus()
    If glbWFC Then
        Dim xCertiNo As String
        If Len(clpBGroup.Text) > 0 Then
            xCertiNo = GetBenCertificateNo(clpBGroup.Text, glbLEE_ID)
            If Len(txtUserText1.Text) = 0 Then
                If Len(txtUserText1.Text) = 0 Then
                    If Not xCertiNo = txtUserText1.Text Then
                        txtUserText1.Text = xCertiNo
                    End If
                End If
            End If
            If IsNumeric(glbBenefitAccount) Then
                txtUserNum1.Text = glbBenefitAccount
            End If
        End If
    End If
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
'City of Kawartha Lakes or City of Timmins or City of Niagara Falls
If (glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W") And Index = 2 Then
    strTmpUnion = clpCode(2).Text
End If

End Sub

Private Function LOA_Document_Found(xEmpnbr, xReason, xFromDate, xToDate) As Long
    Dim rsHRStatus As New ADODB.Recordset
    Dim rsHRStatusDoc As New ADODB.Recordset
    Dim SQLQ As String
    
    LOA_Document_Found = 0
    If Not gsAttachment_DB Then
        Exit Function
    End If
    
    SQLQ = "SELECT * FROM HRSTATUS "
    If IsDate(xFromDate) Or IsDate(xToDate) Then
        SQLQ = SQLQ & " WHERE SC_REASON IN ('LOA') AND SC_EMPNBR=" & xEmpnbr
        SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(xFromDate)
        SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(xToDate)
        SQLQ = SQLQ & " AND SC_NEWEMP='" & xReason & "'"
        rsHRStatus.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRStatus.EOF Then
            SQLQ = "SELECT * FROM HRDOC_HRSTATUS "
            SQLQ = SQLQ & " WHERE SC_DOCKEY = " & rsHRStatus("SC_ID") & " AND SC_EMPNBR=" & xEmpnbr
            SQLQ = SQLQ & " AND SC_TYPE ='LOA'"
            rsHRStatusDoc.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsHRStatusDoc.EOF Then
                LOA_Document_Found = rsHRStatusDoc("SC_DOCKEY")
            Else
                LOA_Document_Found = 0
            End If
            rsHRStatusDoc.Close
            Set rsHRStatusDoc = Nothing
        Else
            LOA_Document_Found = 0
        End If
        rsHRStatus.Close
        Set rsHRStatus = Nothing
    End If
    
End Function

Private Function getLOA_HRSTATUS_ID(xEmpnbr, xReason, xFromDate, xToDate) As Long
    Dim rsHRStatus As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HRSTATUS "
    If IsDate(xFromDate) Or IsDate(xToDate) Then
        SQLQ = SQLQ & " WHERE SC_REASON IN ('LOA') AND SC_EMPNBR=" & xEmpnbr
        SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(xFromDate)
        SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(xToDate)
        SQLQ = SQLQ & " AND SC_NEWEMP='" & xReason & "'"
        rsHRStatus.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRStatus.EOF Then
            getLOA_HRSTATUS_ID = rsHRStatus("SC_ID")
        Else
            getLOA_HRSTATUS_ID = 0
        End If
        rsHRStatus.Close
        Set rsHRStatus = Nothing
    End If
    
End Function

Private Sub UnionDateSet4Vitalaire()
    If NewHireForms.count > 0 Then 'New Hire only
        If clpCode(1).Text = "A" And Left(comEmpType.Text, 1) = "1" And clpPT.Text = "FT" Then
            dlpDate(4).Text = dlpDate(7).Text
        End If
    End If
End Sub


Private Sub clpCode_LostFocus(Index As Integer)
'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" And Index = 2 Then
    If strTmpUnion <> clpCode(2).Text Then
        glbUnionCode = True
        glbTrsVadim1 = ""
        frmNewEmployee.Show 1
        If glbTrsVadim1 <> "Cancel" Then
            txtVadim1.DataField = "ED_VADIM1"
            txtVadim1.Text = glbTrsVadim1
        End If
        glbUnionCode = False
    End If
'City of Timmins
ElseIf glbCompSerial = "S/N - 2375W" And Index = 2 And NewHireForms.count > 0 Then
    If (strTmpUnion <> clpCode(2).Text) And (clpCode(2).Text = "P") Then
        glbUnionCode = True
        glbTrsVadim1 = ""
        frmNewEmployee.lblVadim1.Caption = "PSCAC Rate Level"
        frmNewEmployee.Show 1
        If glbTrsVadim1 <> "Cancel" Then
            txtVadim1.DataField = "ED_VADIM2"
            txtVadim1.Text = glbTrsVadim1
        End If
        glbUnionCode = False
    End If
'City of Niagara Falls
'ElseIf glbCompSerial = "S/N - 2276W" And Index = 2 And NewHireForms.count > 0 Then
'    If (strTmpUnion <> clpCode(2).Text) Then
'        glbUnionCode = True
'        glbTrsVadim1 = ""
'        frmNewEmployee.lblVadim1.Caption = "Union Rate Level"
'        frmNewEmployee.Show 1
'        If glbTrsVadim1 <> "Cancel" Then
'            txtVadim1.DataField = "ED_VADIM2"
'            txtVadim1.Text = glbTrsVadim1
'        End If
'        glbUnionCode = False
'    End If
End If

If Index = 2 Then
    If (glbCompSerial = "S/N - 2385W") Then  'Conservation Halton Ticket #14114
        If clpCode(Index).Text = "CO" Then
            comEmpType.ListIndex = 9
        End If
        If clpCode(Index).Text = "PB" Then
            comEmpType.ListIndex = 4
        End If
    End If
End If

If glbWFC Then 'Ticket #19306
    'If Index = 2 Then
    If Index = 2 Or Index = 1 Then '2-union; 1-status Ticket #20305 Franks 05/17/2011
        Call DispNGSGroups
        Call WFCNGSStartDate 'Ticket #24695 Franks 11/28/2013
    End If
    
    Call WFCSetUnionDate 'Ticket #30376 Franks 07/17/2017
End If

If glbCompSerial = "S/N - 2410W" And Index = 2 Then 'Frontenac  - Ticket #25122 Franks 03/07/2014
    If strTmpUnion <> clpCode(2).Text Then
        Call DispBenGrpFromUnion(clpCode(2).Text)
    End If
End If

'Release 8.1
If Not glbtermopen Then
    Call ShowHide_LOA_Attachment_Buttons
End If

End Sub

Private Sub clpPT_Change()
If (glbCompSerial = "S/N - 2385W") Then  'Conservation Halton Ticket #14114
    If clpPT.Text = "FT" Then
        comEmpType.ListIndex = 1
    End If
    If clpPT.Text = "TR" Then
        comEmpType.ListIndex = 4
    End If
End If
End Sub

Private Sub clpPT_LostFocus()
If glbWFC Then 'Ticket #20712
    Call DispNGSGroups
    If clpPT.Text = "FT" Then Call WFCNGSStartDate 'Ticket #23247 Franks 03/05/2014
End If
End Sub

Private Sub clpVadim1_LostFocus()
    'Ticket #20441
    'If glbWFC Then
    '    Call DispPayGroup(clpVadim1.Text)
    'End If
End Sub
Private Sub clpVadim2_LostFocus() 'Ticket #19266
    'Ticket #20441
    'If glbWFC Then
    '    Call DispNGSSubGroup(clpVadim2.Text)
    'End If
End Sub

Private Sub cmdDemo_Click()
    glbLOAComments = False
    frmEESTATSComm.Show 1
End Sub

Private Sub cmdEditNGSSub_Click()
        glbAccessPswd = False
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            Exit Sub
        End If
        clpVadim1.Enabled = True
        clpVadim1.SetFocus
End Sub

Private Sub cmdEditUserNum1_Click()
        glbAccessPswd = False
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            Exit Sub
        End If
        txtUserNum1.Enabled = True
        txtUserNum1.SetFocus
End Sub

Private Sub cmdEditUserText1_Click()
        glbAccessPswd = False
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            Exit Sub
        End If
        txtUserText1.Enabled = True
        txtUserText1.SetFocus
End Sub

Private Sub cmdEditUserText2_Click()
        glbAccessPswd = False
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            Exit Sub
        End If
        txtUserText2.Enabled = True
        txtUserText2.SetFocus
End Sub

Private Sub cmdEmailImp_Click()
    Dim DgDef, Title$, Msg$, Response%
    
    If Trim(txtFileName.Caption) = "" Then
        MsgBox "File to import not selected. Please select the file to import.", vbExclamation
        cmdEmailImp.SetFocus
        Exit Sub
    ElseIf Dir(txtFileName.Caption) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & txtFileName.Caption & "]", vbExclamation
        cmdEmailImp.SetFocus
        Exit Sub
    Else
        Title$ = "Email Import"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "Are you sure you want to import this Email file?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Sub
        End If
        
        Call Load_Email
                
    End If

End Sub

Private Sub Load_Email()
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim rsEmp As New ADODB.Recordset
    Dim xSkipped As String
    Dim SQLQ As String
    Dim xEmail As String
    Dim xNum As Integer
    Dim xRows As Long
    Dim xRow As Long
    Dim xEmpnbr
    
    
    On Error GoTo Email_Err

    Screen.MousePointer = vbHourglass
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"

    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(txtFileName.Caption)
    Set exSheet = exBook.Worksheets(1)
'    xCols = 1
    xSkipped = ""
    xNum = 0
'    ReDim xTitle(xCols)
'    For X = 1 To xCols
'        xTitle(X) = exSheet.Cells(1, X)
'        Debug.Print "case """ & xTitle(X) & """"
'    Next

    xRows = getRows(exSheet)

    For xRow = 2 To xRows
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
     
        xEmpnbr = exSheet.Cells(xRow, 1)
        xEmail = exSheet.Cells(xRow, 2)
        
        If Not IsNumeric(xEmpnbr) Or xEmpnbr = 0 Or Trim(xEmail) = "" Then
            xSkipped = xSkipped & xEmpnbr & "; "
            xNum = xNum + 1
            If xNum = 10 Then
                xSkipped = xSkipped & vbCrLf
                xNum = 0
            End If
        Else
            Set rsEmp = Nothing
            rsEmp.Open "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR =" & xEmpnbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                rsEmp("ED_EMAIL") = Left(exSheet.Cells(xRow, 2), 60)
                rsEmp.Update
            Else
                xSkipped = xSkipped & xEmpnbr & "; "
                xNum = xNum + 1
                If xNum = 10 Then
                    xSkipped = xSkipped & vbCrLf
                    xNum = 0
                End If
            End If
            rsEmp.Close
            Set rsEmp = Nothing
        End If
    Next
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodPercent = 0
    'MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    

    Screen.MousePointer = vbDefault

    If Len(xSkipped) > 0 Then
        MsgBox "The Email address for the following Employee(s) have been skipped:" & vbCrLf & xSkipped, vbOKOnly + vbInformation, "Import Email Addresses"
    Else
        MsgBox "Employee's Email Addresses have been loaded successfully on Status/Dates screen.", vbOKOnly + vbInformation, "Import Email Addresses"
    End If

Exit Sub

Email_Err:
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    'MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = vbDefault

    If Err.Number = 1004 Then
        MsgBox "Import file not found, try again.", vbOKOnly + vbExclamation, "Email List File Missing"
        Exit Sub
    Else
        MsgBox Err.Description
        Exit Sub
    End If
End Sub

Private Function getRows(exSheet As Object)
Dim x
x = 1
Do While True
    If exSheet.Cells(x, 1) = "" Then
        Exit Do
    Else
        x = x + 1
    End If
Loop
getRows = x - 1
End Function

Private Sub cmdEmailImp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdEmailImpFile_Click()
    glbDocName = "EmailSetup"
    
    AttachmentDialog.DialogTitle = "Select the file to import..."
    AttachmentDialog.Filter = "*.xls;*.xlsx|*.xls;*.xlsx"    '"Word Documents (*.doc;*.docx)|*.doc;*.docx"
    AttachmentDialog.FilterIndex = 1
    AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    AttachmentDialog.ShowOpen
    If Len(AttachmentDialog.FileName) <> 0 Then
        txtFileName.Caption = AttachmentDialog.FileName
    Else
        glbDocName = ""
    End If

End Sub

Private Sub cmdEmailImpFile_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = False
    glbDocName = "Resume"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEESTATS")
End Sub

Private Sub cmdImport1_Click()
    glbDocName = "Termination"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEESTATS")
End Sub

Private Sub cmdImport2_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "LOA"
    If fglbNew Then
        glbDocKey = 0
        'You can only import document for the Save LOA record from Enter a Leave
        MsgBox "You can only Import Document for a saved Leave of Absence record. Please use the Leave and Terminations: 'Enter a Leave' option to record Leave of Absence."
    Else
        glbDocKey = getLOA_HRSTATUS_ID(glbLEE_ID, clpCode(1).Text, dlpDate(15).Text, dlpDate(16).Text)
    
        frmInAttachment.Show 1
        DoEvents
                
        'Update HRDOC_HRSTATUS table with other field values
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update HRDOC_HRSTATUS set SC_STYPE='" & clpCode(1).Text & "', SC_FDATE=" & Date_SQL(dlpDate(15).Text) & " WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR = " & glbLEE_ID & " AND SC_DOCKEY = " & glbDocKey & " AND SC_DOCTYPE = '" & glbDocType & "' AND SC_USRDESC = '" & glbDocDesc & "'"
        gdbAdoIhr001_DOC.CommitTrans
        
        'Call DispimgIcon(Me, "frmETLAY")
        Call ShowHide_LOA_Attachment_Buttons
    End If
End Sub

Private Sub cmdLOAComments_Click()
    Dim xID
    xID = getLOA_HRSTATUS_ID(glbLEE_ID, clpCode(1).Text, dlpDate(15).Text, dlpDate(16).Text)
    If xID <> "" And xID <> 0 Then
        glbLOAComments = True
        frmEESTATSComm.Show 1
    Else
        MsgBox "No Leave of Absence record found for this employee."
    End If
    glbLOAComments = False
End Sub

Private Sub comEmpType_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comEmpType_Click()
'comEmpType.Sorted = True before
If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2384W" Or glbCompSerial = "S/N - 2437W" Then  'St. John's Rehab Hospital 'Ticket #14752
    'St. Marys , KN&V
    If comEmpType.ListIndex <> -1 Then
        txtEmpType.Text = Left(comEmpType.Text, 1)
    End If
ElseIf glbCompSerial = "S/N - 2172W" Then   'Ticket #17077 - County of Lanark
    If comEmpType.ListIndex <> -1 Then
        Select Case comEmpType.ListIndex
            Case 0: txtEmpType.Text = "C"
            Case 1: txtEmpType.Text = "F"
            Case 2: txtEmpType.Text = "P"
            Case 3: txtEmpType.Text = "T"
            Case 4: txtEmpType.Text = "O"
        End Select
    End If
ElseIf glbWFC Then
    If comEmpType.ListIndex <> -1 Then
        Select Case comEmpType.ListIndex
            Case 0: txtEmpType.Text = "Y"
            Case 1: txtEmpType.Text = "N"
        End Select
    End If
ElseIf glbCompSerial = "S/N - 2331W" Then 'Ticket #24402 Cochrane District Franks 09/30/2013
    If comEmpType.ListIndex <> -1 Then
        txtEmpType.Text = comEmpType.ListIndex + 1
    End If
Else
    ' 05/25/2001 Frank Modified code to add "0 - Not Applicable"
    If comEmpType.ListIndex = 0 Then
        txtEmpType.Text = "0"
    ElseIf comEmpType.ListIndex <> -1 Then     ' dkostka - 11/20/2001 - Added comparison to -1 to not fill in if blank.
        If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14736
            Select Case comEmpType.ListIndex
                Case 10: txtEmpType.Text = "A"
                Case 11: txtEmpType.Text = "B"
                Case 12: txtEmpType.Text = "C"  'Ticket #14995
                Case Else
                    txtEmpType.Text = comEmpType.ListIndex
            End Select
        Else
            txtEmpType.Text = comEmpType.ListIndex
        End If
    End If
End If
End Sub

Private Sub ComEType()
comEmpType.Clear
comEmpType.AddItem "0 - Not Applicable"
comEmpType.AddItem "1 - Full Time Salary"
'If glbCompSerial <> "S/N - 2380W" Then   'VitalAire Canada Ticket #14736
    comEmpType.AddItem "2 - Part Time Salary"
'End If
comEmpType.AddItem "3 - Full Time Hourly"
comEmpType.AddItem "4 - Part Time Hourly"
comEmpType.AddItem "5 - Casual/Other"
'Ticket #14736
If glbCompSerial <> "S/N - 2241W" Then  'if not Granite Club
    comEmpType.AddItem "6 - Contract Salary"
    comEmpType.AddItem "7 - Contract Hourly"    '23June99 js
End If
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    'comEmpType.AddItem "8 - 80% Full Time Salary"
    comEmpType.AddItem "8 - 80% @7.5hrs"   'Ticket #14995
ElseIf glbCompSerial = "S/N - 2257W" Then 'HCCAS Ticket #25786 Franks 07/25/2014
    comEmpType.AddItem "8 - Students"
ElseIf glbCompSerial <> "S/N - 2241W" Then  'if not Granite Club - 'Ticket #14736
    comEmpType.AddItem "8 - Salary Pensioners"
End If
'comEmpType.AddItem "9 - Salary Elected officials"
'Added by Bryan 12/08/05 for Granite Club
'If glbCompSerial = "S/N - 2241W" Then  '- Ticket #14736 removed by Hemu
'    comEmpType.AddItem "9 - Commissioned Employees"
'Else
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    comEmpType.AddItem "9 - Former Air Liquide Emp."
ElseIf glbCompSerial <> "S/N - 2241W" Then  'if not Granite Club - Ticket #14736
    comEmpType.AddItem "9 - Salary Elected officials"
End If

'Ticket #14736
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada
    'Ticket #14827
    comEmpType.AddItem "A - 10hr day @86.67"
    comEmpType.AddItem "B - 12 hour day @78"
    comEmpType.AddItem "C - 80% @6hrs"      ''Ticket #14995
    'comEmpType.AddItem "10 - 10hr day @86.67"
    'comEmpType.AddItem "11 - 12 hour day @78"
End If

'Ticket# 10189
If glbCompSerial = "S/N - 2214W" Then 'Casey House Hospice
    comEmpType.Clear
    comEmpType.AddItem "0 - Not Applicable"
    comEmpType.AddItem "1 - Full Time"
    comEmpType.AddItem "2 - Part Time Regular"
    comEmpType.AddItem "3 - Part Time Temporary Full Time"
    comEmpType.AddItem "4 - Part Time Job Share"
    comEmpType.AddItem "5 - Casual Regular"
    comEmpType.AddItem "6 - Casual Temporary Full Time"
End If

'Ticket #14752
If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Hospital
    comEmpType.Clear
    comEmpType.AddItem "A - Temp FT"
    comEmpType.AddItem "B - Temp PT"
    comEmpType.AddItem "C - Casual"
    comEmpType.AddItem "F - FT"
    comEmpType.AddItem "J - Job Share"
    comEmpType.AddItem "P - PT"
    comEmpType.AddItem "S - Student"
    comEmpType.AddItem "X - Terminated"
End If

'Ticket #15794
If glbCompSerial = "S/N - 2390W" Then 'Collectcorp
    comEmpType.Clear
    comEmpType.AddItem "0 - No"
    comEmpType.AddItem "1 - Yes"
End If

'Ticket #16889
If glbCompSerial = "S/N - 2384W" Then 'Town of St. Marys
    comEmpType.Clear
    comEmpType.AddItem "1 - Hourly"
    comEmpType.AddItem "2 - Salary"
    comEmpType.AddItem "3 - Volunteer"
    comEmpType.AddItem "4 - Elected Official"
End If

'Ticket #21096 Franks 09/28/2012
If glbCompSerial = "S/N - 2437W" Then 'KN&V Chartered Accountants LLP
    comEmpType.Clear
    comEmpType.AddItem "1 - Hourly"
    comEmpType.AddItem "2 - Salary"
    comEmpType.AddItem "3 - Volunteer"
    comEmpType.AddItem "4 - Elected Official"
End If

'Ticket #17077
If glbCompSerial = "S/N - 2172W" Then 'County of Lanark
    comEmpType.Clear
    comEmpType.AddItem "C - Part Time Temp"
    comEmpType.AddItem "F - Full Time Regular"
    comEmpType.AddItem "P - Part Time Regular"
    comEmpType.AddItem "T - Full Time Temp"
    comEmpType.AddItem "O - Other" 'Ticket #17076
End If

'Ticket #16395
If glbWFC Then
    comEmpType.Clear
    comEmpType.AddItem "Y - Yes"
    comEmpType.AddItem "N - No"
End If

''Ticket #18602
'If glbCompSerial = "S/N - 2410W" Then 'County of Frontenac
'    comEmpType.Clear
'    comEmpType.AddItem "0 - Not Applicable"
'    comEmpType.AddItem "1 - Full Time Regular"
'    comEmpType.AddItem "2 - Full Time Temp"
'    comEmpType.AddItem "3 - Part Time Regular"
'    comEmpType.AddItem "4 - Part Time Temp(Casual)"
'End If

'Ticket #24402 Franks 09/30/2013
If glbCompSerial = "S/N - 2331W" Then 'Cochrane District
    comEmpType.Clear
    comEmpType.AddItem "1 - FULL"
    comEmpType.AddItem "2 - PART"
    comEmpType.AddItem "3 - SECTEN"
    comEmpType.AddItem "4 - STUDEN"
    comEmpType.AddItem "5 - TEMP"
    comEmpType.AddItem "6 - BOARD"
End If

'Ticket #24864 Franks 06/10/2014
If glbCompSerial = "S/N - 2457W" Then 'McLeod Law
    comEmpType.Clear
    comEmpType.AddItem "S - Salaried Employees"
    comEmpType.AddItem "H - Hourly Employees"
End If
End Sub

Private Function DaysBetween(txtfld1, txtfld2)
Dim datfld1 As Variant, datfld2 As Variant

datfld1 = CVDate(txtfld1)
datfld2 = CVDate(txtfld2)

DaysBetween = DateDiff("d", datfld1, datfld2)

End Function
Function EERetrieve()
Dim SQLQ As String

EERetrieve = False
On Error GoTo EERError

If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
If glbtermopen Then
    If rsDATA2.State <> 0 Then: If rsDATA2.EOF Then rsDATA2.Close Else If rsDATA2.EditMode = adEditAdd Then rsDATA2.CancelUpdate: rsDATA2.Close Else rsDATA2.Close
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

locWHRS = 0
If glbtermopen Then
    SQLQ = "SELECT " & FldList
    SQLQ = SQLQ & " from Term_HREMP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    
    SQLQ = "SELECT * from Term_HRTRMEMP WHERE TERM_SEQ = " & glbTERM_Seq
    rsDATA2.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    If glbOracle Then
        SQLQ = "SELECT Term_DOT,Term_Reason,Term_DOR,Term_Comments, " & Replace(FldList, "TERM_SEQ", "Term_HREMP.TERM_SEQ")
        SQLQ = SQLQ & " FROM Term_HREMP, Term_HRTRMEMP "
        SQLQ = SQLQ & " WHERE Term_HREMP.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ "
        SQLQ = SQLQ & " AND Term_HREMP.TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "SELECT Term_DOT,Term_Reason,Term_DOR,Term_Comments, "
        If glbWFC Then 'Ticket #15248
            SQLQ = SQLQ & "Term_Cause, "
        End If
        SQLQ = SQLQ & Replace(FldList, "TERM_SEQ", "Term_HREMP.TERM_SEQ")
        SQLQ = SQLQ & " FROM Term_HREMP "
        SQLQ = SQLQ & "   INNER JOIN Term_HRTRMEMP ON Term_HREMP.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ"
        SQLQ = SQLQ & " WHERE Term_HREMP.TERM_SEQ = " & glbTERM_Seq
    End If
    
    Data1.RecordSource = SQLQ
Else
    SQLQ = "SELECT " & FldList
    SQLQ = SQLQ & " from HREMP"
    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Data1.RecordSource = SQLQ
    
    If glbWFC Then 'Ticket #25248 Franks 03/24/2014
        locWHRS = GetJHData(glbLEE_ID, "JH_WHRS", 0)
    End If
End If

Data1.Refresh

'--Dates in HREMP_OTHER table - Begin Ticket #15576
If rsDAT_Other.State <> 0 Then: If rsDAT_Other.EOF Then rsDAT_Other.Close Else If rsDAT_Other.EditMode = adEditAdd Then rsDAT_Other.CancelUpdate: rsDAT_Other.Close Else rsDAT_Other.Close
If glbtermopen Then
    If rsDAT_Other.State <> 0 Then: If rsDAT_Other.EOF Then rsDAT_Other.Close Else If rsDAT_Other.EditMode = adEditAdd Then rsDAT_Other.CancelUpdate: rsDAT_Other.Close Else rsDAT_Other.Close
End If

If glbtermopen Then
    DataOther.ConnectionString = glbAdoIHRAUDIT
Else
    DataOther.ConnectionString = glbAdoIHRDB
End If
If glbtermopen Then
    SQLQ = "SELECT " & FldListOther
    SQLQ = SQLQ & " from Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    rsDAT_Other.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    DataOther.RecordSource = SQLQ
Else
    SQLQ = "SELECT " & FldListOther
    SQLQ = SQLQ & " from HREMP_OTHER"
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
    rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    DataOther.RecordSource = SQLQ
End If
If rsDAT_Other.EOF Then
    rsDAT_Other.AddNew
    rsDAT_Other("ER_COMPNO") = "001"
    rsDAT_Other("ER_EMPNBR") = glbLEE_ID
    If glbtermopen Then
        rsDAT_Other("TERM_SEQ") = glbTERM_Seq
    End If
    rsDAT_Other.Update
End If
'Ticket #19932 Franks 03/21/2011
If Not rsDAT_Other.EOF Then
    lblCommDesc.Visible = False
    If Not IsNull(rsDAT_Other("ER_COMMENT")) Then
        If Len(Trim((rsDAT_Other("ER_COMMENT")))) > 0 Then
            lblCommDesc.Visible = True
        End If
    End If
End If
DataOther.Refresh
'--Dates in HREMP_OTHER table - End

Call Display_Value

If Not Data1.Recordset.EOF Then
    If glbCompSerial = "S/N - 2357W" And Data1.Recordset("ED_COUNTRY") = "CANADA" Then   'I.T. Xchange
        lblUnion.FontBold = True
    ElseIf glbCompSerial = "S/N - 2357W" And Data1.Recordset("ED_COUNTRY") <> "CANADA" Then   'I.T. Xchange
        lblUnion.FontBold = False
    End If
End If

EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Resume Next

Exit Function

End Function

'Private Sub DispimgIcon(xTerm)
'Dim SQLQ
'Dim RsTemp As New ADODB.Recordset
'
'If Not gsAttachment_DB Then Exit Sub
'
'lblResume.Visible = True
'If xTerm Then
'    SQLQ = "SELECT * FROM Term_HRDOC_EMP WHERE RE_TYPE='RESUME' AND TERM_SEQ = " & glbTERM_Seq
'Else
'    SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & glbLEE_ID
'End If
''RsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'RsTemp.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
'If Not RsTemp.EOF Then
'    imgSec.Visible = True
'    imgNoSec.Visible = False
'Else
'    imgSec.Visible = False
'    imgNoSec.Visible = True
'End If
'RsTemp.Close
'If gSec_Upd_Basic And Not glbtermopen Then
'    cmdResImp.Visible = True
'Else
'    cmdResImp.Visible = False
'End If
'
'End Sub

Private Sub Command1_Click()
''Dim xMsg
'''frmBENGRLIST.Show 1
''        xMsg = "Does this employee qualify for:"
''        'xMsg = xMsg & " Will this LOA affect the Reporting Authority structures?"
''        frmMsgYesNoUn.lblMsg.Caption = xMsg
''        frmMsgYesNoUn.lblMsg.Alignment = 0
''        Call frmMsgYesNoUn.WFCFrameSetup
''        frmMsgYesNoUn.Show 1
        
End Sub

Private Sub comEmpType_LostFocus()
'Ticket #16395
'Early, Normal and Latest Retirement Dates should be automatically populated.
'If the Pension Eligibility Flag (Employment Type) equals Y (for Yes)
'If NewHireForms.count > 0 Then
'If Type="Y" and these fields are blank, no matter is new hire or not
    If glbWFC Then
        If comEmpType.Text = "Y - Yes" Then
            If IsDate(rsDATA("ED_DOB")) Then
                'Early Retirement Date
                If Len(dlpDate(9).Text) = 0 Then dlpDate(9).Text = WFCPenEmpRetireDate(55, rsDATA("ED_DOB"))
                'Normal  Retirement Date
                If Len(dlpDate(10).Text) = 0 Then dlpDate(10).Text = WFCPenEmpRetireDate(65, rsDATA("ED_DOB"))
                'Early Retirement Date
                If Len(dlpDate(11).Text) = 0 Then dlpDate(11).Text = WFCPenEmpRetireDate(71, rsDATA("ED_DOB"))
            End If
        End If
    End If
'End If
End Sub

Private Sub comUserText1_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comUserText2_Click()
If glbWFC Then 'Ticket #22448 Franks
    If comUserText2.ListIndex <> -1 Then
        txtUserText2.Text = getUserText2(comUserText2.Text)
    End If
'Ticket #24976 - VitalAire Canada Inc.
ElseIf glbCompSerial = "S/N - 2380W" Then
    If comUserText2.ListIndex <> -1 Then
        txtUserText2.Text = Trim(Left(comUserText2.Text, InStr(1, comUserText2.Text, "-") - 1))
    End If
End If
End Sub

Private Function getUserText2(xDesc) 'Ticket #22448 Franks
'format txtUserText2 + " - " to have the description
Dim I As Integer
Dim retVal
    retVal = ""
    I = InStr(1, xDesc, "-")
    If I > 0 Then
        retVal = Trim(Left(xDesc, I - 1))
    End If
    getUserText2 = retVal
End Function

Private Sub comUserText2_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpDate_GotFocus(Index As Integer)
'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" And Index = 2 Then
    dtTmpOMERS = dlpDate(2).Text
End If
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMEESTATS"
If Me.WindowState = 2 Then
    Me.ZOrder BRINGTOFRONT
End If
'Me.cmdModify_Click
Call SET_UP_MODE

If glbWFC Then 'Ticket #22448 Franks
    Call txtUserText2_Change
    If clpCode(1).Enabled Then clpCode(1).SetFocus
End If

End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEESTATS"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim strDOHDate, strFMLADate, strDOHYear
Dim strFMLAYear, strYear, strDecade, strCentury


If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

'cmdPrint.Visible = glbSQL

glbOnTop = "FRMEESTATS"
Screen.MousePointer = HOURGLASS

Call TabStringSetup

Call LabelsSetup

'Ticket #24361 - Remove this option from here and put under Mass Update
'If gSec_Import_Employee Then
'    imgHelp.Visible = True
'    cmdEmailImpFile.Visible = True
'    cmdEmailImp.Visible = True
'    txtFileName.Visible = True
'Else
'    imgHelp.Visible = False
'    cmdEmailImpFile.Visible = False
'    cmdEmailImp.Visible = False
'    txtFileName.Visible = False
'End If

If (glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2409W") Then
    lblUnion.FontBold = True
End If

If glbCompSerial = "S/N - 2297W" Or glbCompSerial = "S/N - 2366W" Then
    lblSen.FontBold = True
End If

If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    lblUserText1.Visible = False
    txtUserText1.Visible = False
End If

If glbCompSerial = "S/N - 2259W" Then
    lblEEType.Visible = False
    comEmpType.Visible = False
    lblLHire.FontBold = True
End If

If glbCompSerial = "S/N - 2394W" Then 'St. John's Ticket #14752
    lblLHire.FontBold = True
    txtUserText2.Visible = False
    txtUserText2.DataField = ""
    dlpDate(17).Visible = True
    dlpDate(17).DataField = "ED_USER_TEXT2"
    lblUserText2.FontBold = True
    lblBen.FontBold = True
End If

'Added by Bryan 12/Jan/06 Ticket#10141, 10188
If glbCompSerial = "S/N - 2378W" Then 'Aurora
    lblODate.FontBold = True
    lblBen.FontBold = True
    lblSalDist.FontBold = True
End If

If glbCompSerial = "S/N - 2453W" Then  'Town of Gander Ticket #24518 Franks 06/04/2015
    lblUnion.FontBold = True
    lblBen.FontBold = True
End If

'Ticket #18844 Franks 01/13/2011  '2383 Town of Orangeville
'Ticket #20666 Franks 07/19/2011  '2429W Municipality of North Perth
'If glbCompSerial = "S/N - 2383W" Or glbCompSerial = "S/N - 2429W" Then
'Ticket #23189 Franks 02/07/2013 - removed this for Orangeville
'2436W Family Day Ticket #24729 01/20/2014
'Ticket #25376 - Community Living Access Support Services
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2436W" Or glbCompSerial = "S/N - 2301W" Then
    lblSalDist.FontBold = True
End If
'If glbCompSerial = "S/N - 2182W" Then lblSalDist.FontBold = True
If glbCompSerial = "S/N - 2466W" Then lblSalDist.FontBold = True 'Chiefs of Ontario Ticket #25879 Franks 09/25/2014

If glbCompSerial = "S/N - 2367W" Then 'Eva's Initiatives
    'lblPenDates.Caption = "RRSP ELIGIBILITY"
    tabDates.Tabs(2).Caption = "RRSP Eligibility"
End If

'Hemu - CollectCorp Inc. - Ticket #14247
If glbCompSerial = "S/N - 2390W" Then
    lblUnion.FontBold = True
End If
If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    lblUnion.FontBold = True
End If

If (glbCompSerial = "S/N - 2394W") Then   ' St. John's Rehab Hospital - Ticket #14572
    lblSalDist.FontBold = True
    lblEEType.Caption = "MediPay Status"
End If

If (glbCompSerial = "S/N - 2390W") Then   'Collectcorp
    lblEEType.Caption = "Pay Source"
End If

'Simona - begin - Assessment Strategies-#14963
If (glbCompSerial = "S/N - 2401W") Then
    lblIPhone.FontBold = True
    lblEmail.FontBold = True
End If
'Simona - end - Assessment Strategies-#14963

'Simona - begin - VitalAire #14995
If (glbCompSerial = "S/N - 2380W") Then
    comEmpType.Visible = False
    clpCode(7).Visible = False
    clpCode(4).Visible = True
    clpCode(4).DataField = "ED_SECTION"
End If
'Simona - end - VitalAire #14995

'Ticket #24996 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    lblSection.Caption = lStr("Section")
    lblSection.Left = lblTitle(20).Left   'Ticket #29230 - Moved the Section to where Location is because we added Union Effective Date. Location is not visible for all & Section is for only Campbell River.
    lblSection.Visible = True
    clpCode(4).Visible = True
    clpCode(4).Left = clpSalDist.Left
    'clpCode(4).Top = txtEmail.Top  'Ticket #29230 - Moved the Section to where Location is (clpCode(0).Top) because we added Union Effective Date. Location is not visible for all & Section is for only Campbell River.
    clpCode(4).Top = clpCode(0).Top
    clpCode(4).Height = 285
    clpCode(4).DataField = "ED_SECTION"
End If


If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #18090
    lblUnion.FontBold = True
    lblSen.FontBold = True
    lblEEType.Visible = False
    comEmpType.Visible = False
    
    'Ticket #22178
    lblSalDist.FontBold = True
    
    'Ticket #20600 Franks 09/01/2011
    Call SamuelFieldsLayout
End If

'Ticket #22710 - County of Perth - Move Vacation Pay Percentage to where Salary Distribution was
If glbCompSerial = "S/N - 2417W" Then
    clpSalDist.Visible = False
    lblSalDist.Caption = "Vacation Pay Percentage"
    medVacPPct.DataField = "ED_VACPC"
    medVacPPct.Left = clpSalDist.Left + 310
    medVacPPct.Top = clpSalDist.Top
    medVacPPct.Visible = True
End If

'Ticket #29617 - Mississaugas of Scugog Island First Nation - Vacation Pay Percentage across from Hire Code
If glbCompSerial = "S/N - 2485W" Then
    clpCode(0).Visible = False
    lblTitle(20).Caption = "Vacation Pay Percentage"
    lblTitle(20).Visible = True
    medVacPPct.DataField = "ED_VACPC"
    medVacPPct.Left = txtUserText2.Left
    medVacPPct.Top = clpCode(0).Top
    medVacPPct.Visible = True
End If

'If glbCompSerial = "S/N - 2391W" Then 'Ticket #26979 Franks 04/27/2015
'    lblSalDist.Visible = False
'    clpSalDist.Visible = False
'End If

'Jerry said this is for everyone
'Ticket #19935 Franks, this is for Samuel only but other customers can see it.
'If Not glbCompSerial = "S/N - 2382W" Then
'    cmdDemo.Enabled = False
'End If

Call TabOrderSetup

'If glbCompSerial = "S/N - 2172W" Then   'County of Lanark
'    lblEEType.Visible = False
'    comEmpType.Visible = False
'End If

'If glbAdv Then  'George Commented for London CCAC changes on Feb 3,2005
'    lblUDay.FontBold = True
'End If
flagFrmLoad = True 'carmen may 00

Call ComEType

'Ticket #23537 - Essex Country Lib. - Remove this logic now
'Ticket #18789 Franks 05/06/2011
'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
'    Call ComEUserText1
'End If

'Ticket #24976 - VitalAire Canada Inc.
If glbCompSerial = "S/N - 2380W" Then
    Call ComEUserTexts
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'Ticket #25040 - Remove the hiding of the Salary Distribution field.
    'lblSalDist.Visible = False
    'clpSalDist.Visible = False
    
    lblUnion.Visible = False
    clpCode(2).Visible = False
End If

If glbVadim Then
    Call VadimControl("Show")
End If

Screen.MousePointer = DEFAULT
If glbCompSerial = "S/N - 2253W" _
    Or glbCompSerial = "S/N - 2252W" _
    Or glbCompSerial = "S/N - 2227W" _
    Or glbCompSerial = "S/N - 2242W" _
    Or glbCompSerial = "S/N - 2269W" _
    Or glbCompSerial = "S/N - 2270W" _
    Or glbCompSerial = "S/N - 2254W" _
    Or glbCompSerial = "S/N - 2214W" Then
        lblODate = "HOOPP Date"
        dlpDate(2).Tag = "00-HOOPP Date"
End If
If glbLinamar Then ' For Linamar
    If NewHireForms.count > 0 Then
        lblDeptStart.FontBold = True
    End If
    lblSen.FontBold = True
    lblSpouse.Visible = True
    chkSpouse.Visible = True
    lblDeptStart.Visible = True
    dlpDate(13).Visible = True
    lblEEType.FontBold = True
    lblTitle(0).Visible = False
    lblTitle(1).Visible = False
    dlpDate(15).Visible = False
    dlpDate(16).Visible = False
    lblBen.Visible = False  'Not available for Linamar
    clpBGroup.Visible = False 'Not available for Linamar
    'lblLang1.Caption = "Primary Language Spoken"
'    lblLang2.Visible = True
    chkElig.Visible = True
    chkElig.DataField = "ED_NOTELIGIBLE"
    lblElig.Top = 380: dlpDate(8).Top = 380: lblAge(0).Top = 380
    lblEarlR.Top = 740: dlpDate(9).Top = 740: lblAge(1).Top = 740
    lblNorR.Top = 1070: dlpDate(10).Top = 1070: lblAge(2).Top = 1070
    lblLateR.Top = 1430: dlpDate(11).Top = 1430: lblAge(3).Top = 1430
    lblHireCode.FontBold = True
    lblElig.FontBold = True
    cmdDemo.Visible = False
    ''Ticket #28846 Franks 07/18/2016 - begin
    'lblWFCMsg.Top = 3350
    'lblWFCMsg.Left = 4500
    'lblWFCMsg.Caption = lStr("Other Date 1") & " is mandatory"
    'lblWFCMsg.Visible = True
    ''Ticket #28846 Franks 07/18/2016 - end
Else
    lblSpouse.Visible = False
    chkSpouse.Visible = False
End If
If glbLambton Then
'    lblSalDist.FontBold = True
'    lblHireCode.FontBold = True
    lblFDay.FontBold = True
End If

'Granite Club
If glbCompSerial = "S/N - 2241W" Then
    lblEEType.FontBold = True
    lblUnion.FontBold = True
    lblSen.FontBold = True
End If

If glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    lblPT.FontBold = False
End If

'Ticket #25469 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    lblUnion.FontBold = True    'Union
    'lblEmail.FontBold = True    'Email     - Don't want Email to be mandatory
End If

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    lblTitle(16).Visible = True
    lblTitle(17).Visible = True
    lblTitle(18).Visible = True
    lblTitle(19).Visible = True
    lblRehired(1).Visible = True
    dlpTermDate.Visible = True
    dlpRehired.Visible = True
    chkRehire.Visible = True
    clpCode(5).Visible = True
    txtComments.Visible = True
    'Ticket #24317 Franks 09/18/2013 - begin
    lblUpdateBy.Visible = True
    lblUserDesc.Visible = True
    lblTitle(3).Visible = True
    dlpDate(35).Visible = True
    'Ticket #24317 Franks 09/18/2013 - end
    If glbWFC Then 'Ticket #15248
        lblTitle(2).Visible = True
        clpCode(3).Visible = True
    End If
End If

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Then
    lblEEType.Caption = "Region"
    Call setCaption(lblEEType)
    comEmpType.Visible = False
    clpCode(7).Visible = True
    clpCode(7).DataField = "ED_REGION"
End If
'Hemu - End

If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    Call WFCUScreenSetup
End If

If glbWFC Then
    'Ticket #19266 - move Vadim fields from Banking screen to Status/Dates
    Call WFCVadimFieldsLayout
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Not rsDATA.EOF Then
    EmpCountry = rsDATA("ED_COUNTRY") & ""
End If

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Then
    If NewHireForms.count > 0 Then
        If Not IsDate(dlpDate(10).Text) Then
            dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
        End If
    End If
End If

'Simona - Begin - Assessment Strategies - Modifications  - Ticket # 14963
If glbCompSerial = "S/N - 2401W" Then
    If NewHireForms.count > 0 Then
        If Not IsDate(dlpDate(9).Text) Then    ''Hemu - Ticket #15753 - Earliest Retirement
            dlpDate(9).Text = DateAdd("yyyy", 50, rsDATA("ED_DOB"))
        End If
        
        If Not IsDate(dlpDate(10).Text) Then    'Normal Retirement
            dlpDate(10).Text = DateAdd("yyyy", 60, rsDATA("ED_DOB"))
        End If
        
        If Not IsDate(dlpDate(10).Text) Then    ''Hemu - Ticket #15753 - Latest/Postponed Retirement
            dlpDate(11).Text = DateAdd("yyyy", 69, rsDATA("ED_DOB"))
        End If
    End If
End If
'Simona - End

'Ticket #26301 - Erb and Erb Insurance Brokers Ltd.
If glbCompSerial = "S/N - 2456W" Then
    If NewHireForms.count > 0 Then
        If Not IsDate(dlpDate(10).Text) Then    'Normal Retirement
            dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
        End If
        If Not IsDate(dlpDate(11).Text) Then    'Latest Retirement
            dlpDate(11).Text = DateAdd("yyyy", 70, rsDATA("ED_DOB"))
        End If
    End If
End If

If glbCountry = "U.S.A." Then   '11Aug99 js
    Label2.Visible = True
    dlpDate(12).Visible = True
Else
    Label2.Visible = False
    dlpDate(12).Visible = False
End If

Screen.MousePointer = HOURGLASS

If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    lblTitle(20).Visible = True
    clpCode(0).Visible = True
    'Ticket #26008 - Other Dates 1 - make it mandatory
    lbOtherDate(0).FontBold = True
End If
If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
    ''Other Dates 1 - make it mandatory
    'lbOtherDate(0).FontBold = True
End If
If glbCompSerial = "S/N - 2410W" Then  'Ticket #18603 - Frontenac
    'Move "Location" to this screen - Jerry
    lblTitle(20).Visible = True
    clpCode(0).Visible = True
    'Ticket #18602
    lblEEType.Visible = False
    comEmpType.Visible = False
    lblBen.FontBold = True
    'Ticket #19071
    lblSalDist.FontBold = True
    lblTitle(20).FontBold = True
    'Ticket #19891
    lblUnion.FontBold = True
End If
If glbCompSerial = "S/N - 2375W" Then  'For Timmis
    lblUnion.FontBold = True
    lblSen.FontBold = True
    lblFDay.FontBold = True
End If
If (glbCompSerial = "S/N - 2357W" And glbEmpCountry = "CANADA") Or glbWFC Then   'I.T. Xchange
    lblUnion.FontBold = True
    If glbWFC Then 'Ticket #30376 Franks 07/17/2017
        lblUnionEDate.FontBold = True
    End If
End If
If glbWFC Then
    txtUserText1.Enabled = False
    'txtUserText2.Enabled = False
    'txtUserNum1.Enabled = False
    cmdEditUserText1.Visible = True
    'cmdEditUserText2.Visible = True
    'cmdEditUserNum1.Visible = True
    'lblTitle(0).FontBold = True
End If
If glbCompSerial = "S/N - 2214W" Then
    lblEmail.FontBold = True
    lblUnion.FontBold = True
    lblTitle(0).FontBold = True
    lblHireCode.FontBold = True
    If clpPT.Text = "FT" Then
        lblIPhone.FontBold = True
    End If
    locHOOPPBen = HOOPPBenFlag(glbLEE_ID)
    If locHOOPPBen Then
        lblODate.FontBold = True
    End If
    comEmpType.Visible = False
End If
If glbCompSerial = "S/N - 2174W" Then 'Kawartha-Haliburton CAS 'Ticket #23382 Franks 04/09/2013
    lblTitle(0).FontBold = True
End If
If (glbCompSerial = "S/N - 2385W") Then ' Conservation Halton 'Ticket #13063
    lblSen.FontBold = True
    'lblUDate.FontBold = True 'Ticket #13165 Union date only needs to be mandatory if ed_pt=FT
End If
If (glbCompSerial = "S/N - 2409W") Then lblSen.FontBold = True 'Ticket #30066 Franks - Skylark Children

If glbCompSerial = "S/N - 2259W" Then 'Test for Oxford Ticket #20892 09/06/2011
    txtTEmpNo.Visible = True
    txtTEmpNames.Visible = True
    txtTEmpNo.Text = lblEENUM.Caption
    txtTEmpNames.Text = lblEEName.Caption
End If

'Ticket #21376 - Charton Hobbs
If glbCompSerial = "S/N - 2418W" Then
    lblEmail.FontBold = True
    lblUserText1.FontBold = True
    lblUDay.FontBold = True
End If

If glbSamuel Then
    'Ticket #21791 Franks 04/09/2012
    If NewHireForms.count > 0 Then
        If GetEmpData(glbLEE_ID, "ED_ADMINBY", "") = "5231" Then
            clpCode(2).Text = "EXEC"
        End If
    End If
End If

If glbCompSerial = "S/N - 2448W" Then 'Workers Health & Safety Centre - Ticket #23428 Franks 03/20/2013
    Call SwitchExpYearUserDefDate
End If

If glbCompSerial = "S/N - 2335W" Then lblUnion.FontBold = True 'Mitchell Plastics Ticket #21866 Franks 04/05/2012
If glbCompSerial = "S/N - 2451W" Then lblUnion.FontBold = True 'Decor Ticket #23848

If glbEntOutStanding$ = "2" Then lblDATE.Caption = dlpDate(7).Text
If glbEntOutStanding$ = "3" Then lblDATE.Caption = dlpDate(6).Text
If glbEntOutStanding$ = "4" Then lblDATE.Caption = dlpDate(5).Text
If glbEntOutStanding$ = "5" Then lblDATE.Caption = dlpDate(3).Text
If glbEntOutStanding$ = "6" Then lblDATE.Caption = dlpDate(4).Text

If glbEntOutStandingS$ = "2" Then lblDateS.Caption = dlpDate(7).Text
If glbEntOutStandingS$ = "3" Then lblDateS.Caption = dlpDate(6).Text
If glbEntOutStandingS$ = "4" Then lblDateS.Caption = dlpDate(5).Text
If glbEntOutStandingS$ = "5" Then lblDateS.Caption = dlpDate(3).Text
If glbEntOutStandingS$ = "6" Then lblDateS.Caption = dlpDate(4).Text

'Ticket #29985 - County of Essex - Autopopulate Union Date with Date of Hire
If glbCompSerial = "S/N - 2192W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(4))) And IsDate(dlpDate(7)) Then
        dlpDate(4).Text = dlpDate(7).Text
    End If
End If

Call INI_Controls(Me)

If glbWFC Then 'Ticket #30376 Franks 07/17/2017
    Call WFCSetUnionDate
End If

cmdModify_Click

Screen.MousePointer = DEFAULT

modSTUPD (False)            '

'If Not gSec_Upd_Basic And Not glbtermopen Then
'    cmdResImp.Visible = False
'End If                          '

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Private Sub TabOrderSetup()
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        clpCode(1).TabIndex = 1
        clpPT.TabIndex = 2
        clpCode(2).TabIndex = 3
        clpSalDist.TabIndex = 4 'Ticket #22066 Franks 05/23/2012
        clpVadim1.TabIndex = 5
        dlpDate(7).TabIndex = 6
        dlpDate(6).TabIndex = 7
        dlpDate(2).TabIndex = 8
        dlpDate(3).TabIndex = 9
    End If
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbWFC And lblWFCMsg.Visible And gSec_Upd_Basic Then  'Ticket #25248 Franks 03/24/2014
    If dlpDate(29).Enabled Then
        Call cmdOK_Click
        Cancel = True
        Exit Sub
    End If
End If
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)



End Sub

Private Sub Form_Resize()
scrHScroll.Width = Me.Width - 120
If Not glbtermopen Then
    'Exit Sub
End If
'If Me.Height >= 4100 + ScrFrame.Height + 700 Then
If Me.Height >= 10560 Then
    scrControl.Value = 0
    ScrFrame.Top = 4200
    scrControl.Visible = False
    'Exit Sub
Else
    scrControl.Visible = True
    'scrControl.Max = 3645
    scrControl.Max = 2400
    scrControl.Left = Me.Width - scrControl.Width - 120
    If Me.Height - scrControl.Top - panControls.Height - 400 > 0 Then
        scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
    End If
End If

'Horizontal Scroll
ScrFrame.Height = Me.ScaleHeight - (panEEDESC.Height + scrHScroll.Height + 200)
If Me.Width >= 12705 Then
    scrHScroll.Value = 0
    scrHScroll.Visible = False
Else
    scrHScroll.Visible = True
    scrHScroll.Top = 100
    scrHScroll.Width = Me.Width - 120
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEESTATS = Nothing
    Call NextForm
End Sub

Public Sub imgEmail_Click()
    Call txtEmail_DblClick
End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "Import File must be an Excel Spreadsheet with the following format: "
    MsgStr = MsgStr & Chr(10) & "        1. First row can be a Header row."
    MsgStr = MsgStr & Chr(10) & "        2. Data to import should start from 2nd row."
    MsgStr = MsgStr & Chr(10) & "        3. Column order to Import:"
    MsgStr = MsgStr & Chr(10) & vbTab & "a. Column 1: Employee #"
    MsgStr = MsgStr & Chr(10) & vbTab & "b. Column 2: Email Address"
    MsgBox MsgStr, vbInformation, "info:HR - Import File Format"
End Sub

Private Sub imgSec_Click()
'If glbtermopen Then
'    Call FillMemoFile(glbTERM_Seq)
'Else
'    Call FillMemoFile(lblEEID)
'End If
    Dim SQLQ
    glbDocName = "Resume"
    SQLQ = getSQL("frmEESTATS")
    Call FillMemoFile(SQLQ, "Resume")
End Sub

Private Sub imgSec1_Click()
    Dim SQLQ
    glbDocName = "Termination"
    SQLQ = getSQL("frmEESTATS")
    Call FillMemoFile(SQLQ, "Termination")
End Sub

Private Sub imgSec2_Click()
    Dim SQLQ
    glbDocName = "LOA"
    glbDocKey = getLOA_HRSTATUS_ID(glbLEE_ID, clpCode(1).Text, dlpDate(15).Text, dlpDate(16).Text)
    SQLQ = getSQL("frmETLAY")
    Call FillMemoFile(SQLQ, "LOA")
End Sub

'Private Function FillMemoFile(zEMPNBR) ' As Long)
'    On Error GoTo ErrHandler:
'    Dim rsPHOTO As New ADODB.Recordset
'    Dim byteChunk() As Byte
'
'    Dim Offset As Long
'    Dim Totalsize As Long
'    Dim Remainder As Long
'
'    Dim FieldSize As Long
'    Dim FileNumber As Integer
'    Const HeaderSize As Long = 78
'    Const ChunkSize As Long = 100
'    Dim TempFile As String
'    Dim TempDir As String * 255
'    Dim FileExt As String
'    Dim SQLQ
'    GetTempPath 255, TempDir
'    'TempFile = Replace(Replace(TempDir, Chr(0), "") & "\tempfile.tmp", "\\", "\")
'
'    If zEMPNBR = 0 Then Exit Function
'    If glbtermopen Then
'        SQLQ = "select * from Term_HRDOC_EMP WHERE RE_TYPE='RESUME' AND TERM_SEQ = " & zEMPNBR
'    Else
'        SQLQ = "select * from HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & zEMPNBR
'    End If
'    'rsPHOTO.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    rsPHOTO.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic, adLockOptimistic
'
'    If rsPHOTO.EOF Then Exit Function
'    If IsNull(rsPHOTO("RE_FILEEXT")) Then
'        FileExt = ""
'        TempFile = Replace(Replace(TempDir, Chr(0), "") & "\IhrResume.tmp", "\\", "\")
'    Else
'        FileExt = rsPHOTO("RE_FILEEXT")
'        TempFile = Replace(Replace(TempDir, Chr(0), "") & "\IhrResume." & FileExt & "", "\\", "\")
'    End If
'
'
'    FileNumber = FreeFile
'    If (Dir(TempFile)) <> "" Then
'        Call FileAttribute(TempFile, "-r", TempDir)
'        Call Pause(2)
'        Kill TempFile
'    End If
'    Open TempFile For Binary Access Write As FileNumber
'
'    ReDim byteChunk(rsPHOTO("RE_DOC").ActualSize)
'    byteChunk() = rsPHOTO("RE_DOC").GetChunk(rsPHOTO("RE_DOC").ActualSize)
'    Put FileNumber, , byteChunk()
'
'    Close FileNumber
'    'Kill (TempFile)
'    rsPHOTO.Close
'
'    'Read only
'    Call FileAttribute(TempFile, "+r", TempDir)
''    TempFile2 = Replace(Replace(TempDir, Chr(0), "") & "\IhrDoc.Bat", "\\", "\")
''    FileNumber = FreeFile
''    Open TempFile2 For Output As #5
''    Print #5, "attrib +r " & GetShortName(TempFile)
''    Close #5
''    Shell "cmd /c " & GetShortName(TempFile2)
'
'    'Open the attachment
'    Shell "cmd /c " & GetShortName(TempFile)
'
'    Exit Function
'
'ErrHandler:
'    MsgBox Err.Description, , "Error "
'
'End Function
'Private Sub FileAttribute(xFileName, xAttribute, xTempDir)
'Dim TempFile2
'    TempFile2 = Replace(Replace(xTempDir, Chr(0), "") & "\IhrDoc.Bat", "\\", "\")
'    Open TempFile2 For Output As #5
'    Print #5, "attrib " & xAttribute & " " & GetShortName(xFileName)
'    Close #5
'    Shell "cmd /c " & GetShortName(TempFile2)
'End Sub
Private Sub lblEEID_Change()

If Len(txtSurname.Text) > 0 And Len(txtFName.Text) > 0 Then  ' don't do on add new until in
    frmEESTATS.Caption = "Status/Dates - " & Left$(txtSurname, 5)
    frmEESTATS.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If
lblEENUM = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub modSTUPD(YN)
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

    
clpPT.Enabled = TF
If glbCompSerial = "S/N - 2380W" Then
    'VitalAire Ticket #12142, user only can change this on Section field on Demo screen
    comEmpType.Enabled = False
ElseIf glbCompSerial = "S/N - 2172W" Then 'Ticket #19588 - County of Lanark
    'user needs both employee and bank maintain right
    comEmpType.Enabled = (TF And gSec_Upd_Banking)
Else
    comEmpType.Enabled = TF
End If
clpCode(1).Enabled = TF
If glbCompSerial = "S/N - 2172W" Then 'Ticket #19588 - County of Lanark
    'user needs both employee and bank maintain right
    clpCode(2).Enabled = (TF And gSec_Upd_Banking)
Else
    clpCode(2).Enabled = TF
End If
'clpCode(3).Enabled = TF
'clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
clpCode(6).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
dlpDate(2).Enabled = TF
dlpDate(3).Enabled = TF
dlpDate(4).Enabled = TF
dlpDate(5).Enabled = TF
dlpDate(6).Enabled = TF
dlpDate(7).Enabled = TF
dlpDate(8).Enabled = TF
dlpDate(9).Enabled = TF
dlpDate(10).Enabled = TF
dlpDate(11).Enabled = TF
txtExpYear.Enabled = TF
dlpDate(14).Enabled = TF
dlpDate(13).Enabled = TF
dlpDate(15).Enabled = TF
dlpDate(16).Enabled = TF
dlpDate(34).Enabled = TF
clpSalDist.Enabled = TF
clpBGroup.Enabled = TF
txtUserText1.Enabled = TF
txtUserText2.Enabled = TF
txtUserNum1.Enabled = TF

If Label2.Visible = True Then
    dlpDate(12).Enabled = TF
End If

txtEmail.Enabled = TF
txtIPHONE.Enabled = TF
chkSpouse.Enabled = TF
If txtUserNum2.Visible Then
    txtUserNum2.Enabled = TF
End If
If clpVadim2.Visible Then
    clpVadim2.Enabled = TF
End If
If glbtermopen Then
    dlpRehired.Enabled = TF
    dlpTermDate.Enabled = TF
    clpCode(5).Enabled = TF
    txtComments.Enabled = TF
End If
If glbLinamar Then
    If clpCode(1).Text = "TEMP" Then
        dlpDate(4).Enabled = False
        dlpDate(3).Enabled = False
    End If
End If
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/28/2014
    TF = getFamilyDayUpdateRight(UpdateRight, glbLEE_ID)
    If Not TF Then
        clpSalDist.Enabled = False
        txtUserText2.Enabled = False
    End If
End If

End Sub

Private Sub medVacPPct_GotFocus()
Call SetPanHelp(Me.ActiveControl)
If Len(medVacPPct) > 0 Then
    medVacPPct = medVacPPct * 100
End If
End Sub

Private Sub medVacPPct_KeyPress(KeyAscii As Integer)
If Not IsNumericEntry(KeyAscii) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medVacPPct_LostFocus()
If (Not IsNumeric(medVacPPct)) And medVacPPct.DataChanged Then medVacPPct = 0
If Len(medVacPPct) > 0 Then
    medVacPPct = medVacPPct / 100
End If
End Sub

Private Sub scrControl_Change()
'ScrFrame.Top = 4100 - scrControl.Value
ScrFrame.Top = 4080 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
ScrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
TopFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub



Private Sub tabDates_Click()
    If tabDates.SelectedItem.Key = "keyEmpDate" Then
        fraDateEmp.Visible = True
        fraDatePension.Visible = False
        fraDateOther.Visible = False
    End If
    If tabDates.SelectedItem.Key = "keyPenDate" Then
        fraDatePension.Visible = gSec_Show_DOB  'True   - Ticket #17919 - Do not show if user do not have access to BirthDate
        fraDateEmp.Visible = False
        fraDateOther.Visible = False
    End If
    If tabDates.SelectedItem.Key = "keyOtherDate" Then
        fraDateOther.Visible = True
        fraDatePension.Visible = False
        fraDateEmp.Visible = False
    End If
End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

'Hemu - 07/23/2003 Begin - Sending email
Private Sub txtEmail_DblClick()
On Error GoTo Email_Err
    If gsEMAIL_SENDING Then
        If Len(txtEmail.Text) > 0 Then
            frmSendEmail.txtTo.Text = txtEmail.Text
            frmSendEmail.Tag = ""
            frmSendEmail.Show 1
        Else
            MsgBox "Email Address is blank."
        End If
    End If
    Exit Sub
    
Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next
    
End Sub
'Hemu - 07/23/2003 End


Private Sub txtExpYear_Change()
Dim xYear
If IsNumeric(txtExpYear) And Val(txtExpYear) > 0 Then
    xYear = Year(Date) - Val(txtExpYear)
    lblYear(2) = xYear & IIf(xYear > 1, " Years", " Year")
Else
    lblYear(2) = ""
End If
End Sub

Private Sub txtExpYear_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpDate_Change(Index As Integer)
Dim xYear, xDATE
Dim strSQL As String
If flagFrmLoad = False Then Exit Sub 'carmen may 00

If Index > 5 And Index < 8 Then
    lblYear(Index - 6).Visible = False
    If Not IsDate(dlpDate(Index).Text) Then Exit Sub
    lblYear(Index - 6).Visible = True
    If glbtermopen Then
        If IsDate(dlpTermDate.Text) Then
            xYear = DateDiff("d", CVDate(dlpDate(Index).Text), dlpTermDate.Text)
        Else
            xYear = 0
        End If
    Else
        xYear = DateDiff("d", CVDate(dlpDate(Index).Text), Date)
    End If
    xYear = Round(xYear / 365, 1)
    lblYear(Index - 6) = xYear & IIf(xYear <> 1, " Years", " Year")
End If

       
If Index > 7 And Index < 12 Then
    lblAge(Index - 8).Visible = False
    If Not IsDate(dlpDate(Index).Text) Then Exit Sub
    lblAge(Index - 8).Visible = True
    xDATE = CVDate(dlpDate(Index).Text)
    xYear = DateDiff("d", rsDATA("ED_DOB"), xDATE)
    xYear = Round(xYear / 365, 1)
    lblAge(Index - 8) = "At age: " & xYear
End If

If Index > 17 And Index < 24 Then
    lblAge(Index - 14).Visible = False
    If Not IsDate(dlpDate(Index).Text) Then Exit Sub
    lblAge(Index - 14).Visible = True
    xDATE = CVDate(dlpDate(Index).Text)
    xYear = DateDiff("d", rsDATA("ED_DOB"), xDATE)
    xYear = Round(xYear / 365, 1)
    lblAge(Index - 14) = "At age: " & xYear
End If

End Sub

Private Sub dlpDate_LostFocus(Index As Integer)
'ticket #8923
'If glbLinamar And Index = 7 And IsDate(dlpDate(7).Text) Then
'    dlpDate(8).Text = DateAdd("yyyy", 2, dlpDate(7).Text)
'End If
'If IsDate(dlpDate(Index).Text) Then
'    dlpDate(Index).Text = CVDate(dlpDate(Index).Text)
'End If

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(5))) And IsDate(dlpDate(7)) Then
        dlpDate(5).Text = dlpDate(7).Text
    End If
End If
'Hemu - End
If glbCompSerial = "S/N - 2259W" And NewHireForms.count > 0 Then 'County of Essex new hire
    If (Not IsDate(dlpDate(5))) And IsDate(dlpDate(7)) Then
        dlpDate(5).Text = dlpDate(7).Text
    End If
End If

'Simona - begin - Assessment Strategies May 27,2008 - if ED_SENDTE is null then seniority date equals date of hire
If glbCompSerial = "S/N - 2401W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(6))) And IsDate(dlpDate(7)) Then
        dlpDate(6).Text = dlpDate(7).Text
    End If
End If
'Simona - end - Assessment Strategies May 27,2008

'Ticket #18090 Frank 03/04/2010 Samuel
If glbCompSerial = "S/N - 2382W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(6))) And IsDate(dlpDate(7)) Then
        dlpDate(6).Text = dlpDate(7).Text
    End If
End If

'Hemu - Autopopulate Seniority Date and User Defined Date with Date of Hire for CollectCorp Inc.
If glbCompSerial = "S/N - 2390W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(6))) And IsDate(dlpDate(7)) Then
        dlpDate(6).Text = dlpDate(7).Text
    End If
    If (Not IsDate(dlpDate(3))) And IsDate(dlpDate(7)) Then
        dlpDate(3).Text = dlpDate(7).Text
    End If
End If

'Ticket #20396 - BACI
If glbCompSerial = "S/N - 2431W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(6))) And IsDate(dlpDate(7)) Then
        dlpDate(6).Text = dlpDate(7).Text
    End If
End If

If Index = 7 Then Call FMLA       'jaddy oct 4,99
If glbLambton Then 'Ticket# 6355
    If Index = 7 Then
        If Len(dlpDate(0)) = 0 Then
            dlpDate(0) = dlpDate(7)
        End If
        If Len(dlpDate(3)) = 0 And UCase(clpCode(1)) = "PERM" And UCase(clpPT) = "FT" Then
            dlpDate(3) = dlpDate(7)
        End If
    End If
End If
If NewHireForms.count > 0 Then 'New Hire only
    If IsDate(dlpDate(7).Text) Then
        If Not IsDate(dlpDate(15).Text) Then
            dlpDate(15).Text = dlpDate(7).Text
        End If
        'City of Timmins
        If glbCompSerial = "S/N - 2375W" Then
            If IsDate(dlpDate(2).Text) Then     'If OMERS Date available then get Retirement date
                If Not IsDate(dlpDate(10).Text) Then 'Retire at 65 years old
                    dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
                End If
            Else
                dlpDate(10).Text = ""
            End If
        ElseIf glbWFC Then 'Ticket #24695 Franks 11/28/2013
            If Not IsDate(dlpDate(10).Text) Then 'Retire at 65 years old
                dlpDate(10).Text = getWFCRetireDate(rsDATA("ED_DOB"))
            End If
        Else
            If Not IsDate(dlpDate(10).Text) Then 'Retire at 65 years old
                dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
            End If
        End If
    End If
    If glbWFC Then 'Ticket #19266 Franks 12/13/2010
        If Index = 7 Then
            Call WFC_NGS_Trans
        End If
    End If
    
    'Ticket #26301 - Erb and Erb Insurance Brokers Ltd.
    If glbCompSerial = "S/N - 2456W" Then
        If NewHireForms.count > 0 Then
            If Not IsDate(dlpDate(10).Text) Then    'Normal Retirement
                dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
            End If
            If Not IsDate(dlpDate(11).Text) Then    'Latest Retirement
                dlpDate(11).Text = DateAdd("yyyy", 70, rsDATA("ED_DOB"))
            End If
        End If
    End If
    
Else
    'City of Timmins
    If glbCompSerial = "S/N - 2375W" Then
        If IsDate(dlpDate(2).Text) Then     'If OMERS Date available then get Retirement date
            If Not IsDate(dlpDate(10).Text) Then 'Retire at 65 years old
                dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
            End If
        Else
            dlpDate(10).Text = ""
        End If
    End If
    If glbWFC And glbtermopen Then 'Ticket #23948 Franks 06/24/2013
        If Index = 21 Then
            If IsDate(dlpDate(21).Text) Then dlpTermDate.Text = dlpDate(21).Text
        End If
    End If
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" And Index = 2 Then
    If dtTmpOMERS <> dlpDate(2).Text Then
        glbOMERS_Date = True
        glbTrsDIV = ""
        frmNewEmployee.Show 1
        If glbTrsDIV <> "Cancel" Then
            txtRPP.DataField = "ED_PENSION"
            txtRPP.Text = glbTrsDIV
        End If
        glbOMERS_Date = False
    End If
End If

'City of Timmins
If glbCompSerial = "S/N - 2375W" And Index = 2 Then
    txtRPP.DataField = "ED_PENSION"
    If Len(dlpDate(2)) > 0 Then
        If clpCode(2) = "D" Or clpCode(2) = "E" Or clpCode(2) = "G" Then    'Ticket #15756 - Police or Fire
            txtRPP.Text = "2"
        Else
            txtRPP.Text = "1"
        End If
        'dlpDate(10) = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
    Else
        txtRPP.Text = "Null"
        dlpDate(10) = ""
    End If
End If

If (glbCompSerial = "S/N - 2385W") Then ' Conservation Halton 'Ticket #13063
    If Index = 7 Then
        If NewHireForms.count > 0 Then 'New Hire only
            If Len(dlpDate(6).Text) = 0 Then
                dlpDate(6).Text = dlpDate(7).Text
            End If
            If Len(dlpDate(4).Text) = 0 Then
                If clpPT.Text = "FT" Then
                    dlpDate(4).Text = dlpDate(7).Text
                End If
            End If
        End If
    End If
End If

'Ticket #14217 VitalAire
If glbCompSerial = "S/N - 2380W" Then
    Call UnionDateSet4Vitalaire
End If

If glbCompSerial = "S/N - 2214W" And NewHireForms.count > 0 Then   'Ticket #14550
    If Not IsDate(dlpDate(0)) Then
        dlpDate(0).Text = dlpDate(7).Text
    End If
    If Not IsDate(dlpDate(6)) Then
        dlpDate(6).Text = dlpDate(7).Text
    End If
    If Not IsDate(dlpDate(2)) Then
        dlpDate(2).Text = dlpDate(7).Text
    End If
End If

If glbCompSerial = "S/N - 2408W" And NewHireForms.count > 0 Then 'Township of Wilmot - Ticket #15785
    If (Not IsDate(dlpDate(5))) And IsDate(dlpDate(7)) Then
        dlpDate(5).Text = dlpDate(7).Text
    End If
    If (Not IsDate(dlpDate(4))) And IsDate(dlpDate(7)) And clpPT.Text = "FT" Then
        dlpDate(4).Text = dlpDate(7).Text
    End If
End If

'Granite Club - Ticket #19019
'Autopopulate User Defined Date with Date of Hire
If glbCompSerial = "S/N - 2241W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(3))) And IsDate(dlpDate(7)) Then
        dlpDate(3).Text = dlpDate(7).Text
    End If
End If

'County of Lanark - Ticket #21199
If glbCompSerial = "S/N - 2172W" Then
    If IsDate(dlpDate(2).Text) Then     'If OMERS Date entered then calculate Normal & Earliest Retirement date
        If Not IsDate(dlpDate(9).Text) Then 'Earliest Retirement at 65 years old
            dlpDate(9).Text = DateAdd("yyyy", 55, rsDATA("ED_DOB"))
        End If
        If Not IsDate(dlpDate(10).Text) Then 'Normal Retirement at 65 years old
            dlpDate(10).Text = DateAdd("yyyy", 65, rsDATA("ED_DOB"))
        End If
    Else
        dlpDate(9).Text = ""
        dlpDate(10).Text = ""
    End If
End If

'Ticket #29985 - County of Essex - Autopopulate Union Date with Date of Hire
If glbCompSerial = "S/N - 2192W" And NewHireForms.count > 0 Then
    If (Not IsDate(dlpDate(4))) And IsDate(dlpDate(7)) Then
        dlpDate(4).Text = dlpDate(7).Text
    End If
End If

If glbWFC Then 'Ticket #25248 Franks 03/24/2014
    If Index = 29 Then
        Call chkWFCOtherDate6("Disp")
    End If
End If
    
End Sub

Private Sub txtEmail_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtEmpType_Change()
If flagFrmLoad = False Then Exit Sub 'carmen may 00
comEmpType.ListIndex = -1
If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Hospital 'Ticket #14752
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
ElseIf glbCompSerial = "S/N - 2384W" Then 'Ticket #16978 St. Marys
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
ElseIf glbCompSerial = "S/N - 2437W" Then 'Ticket #21096 KN&V
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
ElseIf glbCompSerial = "S/N - 2172W" Then 'Ticket #17077 - County of Lanark
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
ElseIf glbCompSerial = "S/N - 2331W" Then 'Ticket #24402 Cochrane District Franks 09/30/2013
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
ElseIf glbWFC Then
    comEmpType.ListIndex = GetEmpTypeIndex(txtEmpType.Text)
Else
    If Val(txtEmpType) > 0 And Val(txtEmpType) <= 9 And glbCompSerial <> "S/N - 2214W" Then
        comEmpType.ListIndex = Val(txtEmpType)
    ElseIf (glbCompSerial = "S/N - 2214W" And Val(txtEmpType) > 0 And Val(txtEmpType) <= 6) Then
        comEmpType.ListIndex = Val(txtEmpType)
    ElseIf glbCompSerial = "S/N - 2214W" And Val(txtEmpType) > 6 Then
        comEmpType.ListIndex = 0
    Else
        If txtEmpType = "0" Then
            comEmpType.ListIndex = 0
        ElseIf glbCompSerial = "S/N - 2380W" Then   'Ticket #14827
            Select Case txtEmpType
                Case "A": comEmpType.ListIndex = 10
                Case "B": comEmpType.ListIndex = 11
                Case "C": comEmpType.ListIndex = 12 'Ticket #14995
            End Select
        End If
    End If
End If
End Sub

Private Function GetEmpTypeIndex(xEmpType)
Dim xIndex As Integer
    xIndex = -1
    If glbCompSerial = "S/N - 2394W" Then 'St. Johns
        Select Case xEmpType
        Case "A": xIndex = 0
        Case "B": xIndex = 1
        Case "C": xIndex = 2
        Case "F": xIndex = 3
        Case "J": xIndex = 4
        Case "P": xIndex = 5
        Case "S": xIndex = 6
        Case "X": xIndex = 7
        End Select
    End If
    If glbCompSerial = "S/N - 2384W" Then 'St. Marys
        Select Case xEmpType
        Case "1": xIndex = 0
        Case "2": xIndex = 1
        Case "3": xIndex = 2
        Case "4": xIndex = 3
        End Select
    End If
    If glbCompSerial = "S/N - 2437W" Then 'Ticket #21096 KN&V
        Select Case xEmpType
        Case "1": xIndex = 0
        Case "2": xIndex = 1
        Case "3": xIndex = 2
        Case "4": xIndex = 3
        End Select
    End If
    If glbCompSerial = "S/N - 2172W" Then 'County of Lanark
        Select Case xEmpType
        Case "C": xIndex = 0
        Case "F": xIndex = 1
        Case "P": xIndex = 2
        Case "T": xIndex = 3
        Case "O": xIndex = 4
        End Select
    End If
    If glbWFC Then
        Select Case xEmpType
        Case "Y": xIndex = 0
        Case "N": xIndex = 1
        End Select
    End If
    If glbCompSerial = "S/N - 2331W" Then 'Ticket #24402 Cochrane District Franks 09/30/2013
        Select Case xEmpType
        Case "1": xIndex = 0
        Case "2": xIndex = 1
        Case "3": xIndex = 2
        Case "4": xIndex = 3
        Case "5": xIndex = 4
        Case "6": xIndex = 5
        End Select
    End If
    
    GetEmpTypeIndex = xIndex
End Function

Private Sub txtIPHONE_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub UpdEICodeForGraniteClub()
    If glbCompSerial <> "S/N - 2241W" Then Exit Sub
    If SavPT = "" Then
        If clpPT = "FT" Then
            rsDATA("ED_UIC") = Null
        Else
            rsDATA("ED_UIC") = "A"
        End If
    ElseIf SavPT <> clpPT And (SavPT = "FT" Or clpPT = "FT") Then
        MsgBox "Please check the employee's EI Code and make a change if necessary.", vbOKOnly, "Check EI code"
    End If

End Sub

Private Function UpdEntOut()
Dim rsAT As New ADODB.Recordset
Dim xWDate, xWDateS, Msg, Response, x1, x2
Dim SQLQ, xVAC, xSICK, xlen, xLapYr
Dim xtVACT, xtEFDATE, xtETDATE

UpdEntOut = True
xWDate = ""
xWDateS = ""
x1 = ""
x2 = ""

If glbEntOutStanding$ = "1" And glbEntOutStandingS$ = "1" Then
    Exit Function
End If
If NewHireForms.count > 0 Then Exit Function

If glbEntOutStanding$ = "2" Then
    If dlpDate(7).Text <> lblDATE.Caption Then xWDate = dlpDate(7).Text
End If

If glbEntOutStanding$ = "3" Then
    If dlpDate(6).Text <> lblDATE.Caption Then xWDate = dlpDate(6).Text
End If

If glbEntOutStanding$ = "4" Then
    If dlpDate(5).Text <> lblDATE.Caption Then xWDate = dlpDate(5).Text
End If

If glbEntOutStanding$ = "5" Then
    If dlpDate(3).Text <> lblDATE.Caption Then xWDate = dlpDate(3).Text
End If

If glbEntOutStanding$ = "6" Then
    If dlpDate(4).Text <> lblDATE.Caption Then xWDate = dlpDate(4).Text
End If

If glbEntOutStandingS$ = "2" Then
    If dlpDate(7).Text <> lblDateS.Caption Then xWDateS = dlpDate(7).Text
End If

If glbEntOutStandingS$ = "3" Then
    If dlpDate(6).Text <> lblDateS.Caption Then xWDateS = dlpDate(6).Text
End If

If glbEntOutStandingS$ = "4" Then
    If dlpDate(5).Text <> lblDateS.Caption Then xWDateS = dlpDate(5).Text
End If

If glbEntOutStandingS$ = "5" Then
    If dlpDate(3).Text <> lblDateS.Caption Then xWDateS = dlpDate(3).Text
End If

If glbEntOutStandingS$ = "6" Then
    If dlpDate(4).Text <> lblDateS.Caption Then xWDateS = dlpDate(4).Text
End If

If xWDate = "" And xWDateS = "" Then Exit Function

If xWDate <> "" Then
    Select Case glbEntOutStanding$
    Case "2": x1 = lStr("Original Hire Date")
    Case "3": x1 = lStr("Seniority Date")
    Case "4": x1 = lStr("Last Hire Date")
    Case "5": x1 = lStr("User Defined Date")
    Case "6": x1 = "Union Date "
    Case Else: x1 = "Entitlements Date "
    End Select
End If

If xWDateS <> "" Then
    Select Case glbEntOutStandingS$
    Case "2": x2 = lStr("Original Hire Date")
    Case "3": x2 = lStr("Seniority Date")
    Case "4": x2 = lStr("Last Hire Date")
    Case "5": x2 = lStr("User Defined Date")
    Case "6": x2 = "Union Date "
    Case Else: x2 = "Entitlements Date "
    End Select
End If

If x1 = "" Then x1 = x2
If x1 = x2 Then x2 = ""

Msg = "Change in " & Chr(34)
If Len(x1) > 1 Then Msg = Msg & x1
If Len(x2) > 1 Then Msg = Msg & " and " & x2
Msg = Msg & Chr(34) & " will affect employees outstanding "
Msg = Msg & "sick and vacation entitlements."
Msg = Msg & Chr(10) & "Do you wish to proceed and recalculate the "
Msg = Msg & "Employee's outstanding entitlement ?"
Response = MsgBox(Msg, 52, "Warning")

If Response = IDNO Then
    UpdEntOut = False
    
    Select Case glbEntOutStandingS$
    Case "2": dlpDate(7).SetFocus
    Case "3": dlpDate(6).SetFocus
    Case "4": dlpDate(5).SetFocus
    Case "5": dlpDate(3).SetFocus
    Case "6": dlpDate(4).SetFocus
    Case Else: dlpDate(7).SetFocus
    End Select
    
    Select Case glbEntOutStanding$
    Case "2": dlpDate(7).SetFocus
    Case "3": dlpDate(6).SetFocus
    Case "4": dlpDate(5).SetFocus
    Case "5": dlpDate(3).SetFocus
    Case "6": dlpDate(4).SetFocus
    Case Else: dlpDate(7).SetFocus
    End Select
    
    Exit Function
End If

' OK to Update Entitlements
UpdEntOut = True
xtVACT = Val(txtVACT)
txtVACT = 0
txtSICKT = 0
xtEFDATE = txtfDate
xtETDATE = txtTDate

If Year(Now) / 4 = Int(Year(Now) / 4) Then
    xLapYr = True
Else
    xLapYr = False
End If

If xWDate <> "" Then
    If IsDate(xWDate) Then
        If month(xWDate) = 2 And Day(xWDate) = 29 And Not xLapYr Then
            xWDate = DateAdd("d", -1, xWDate)
        End If
        xlen = DateDiff("yyyy", CVDate(xWDate), Now)
        xWDate = DateAdd("yyyy", xlen, xWDate)
        If DateValue(xWDate) > Now Then
            xlen = InStr(4, xWDate, "/")
            xWDate = DateAdd("yyyy", -1, xWDate)
        End If
        txtfDate = xWDate
        xWDate = DateAdd("yyyy", 1, xWDate)
        xWDate = DateAdd("d", -1, xWDate)
        txtTDate = xWDate
    Else
        txtfDate = ""
        txtTDate = ""
        txtVACT = 0
    End If
End If
If xWDateS <> "" Then
    If IsDate(xWDateS) Then
        If month(xWDateS) = 2 And Day(xWDateS) = 29 And Not xLapYr Then
            xWDateS = DateAdd("d", -1, xWDateS)
        End If
        xlen = DateDiff("yyyy", CVDate(xWDateS), Now)
        xWDateS = DateAdd("yyyy", xlen, xWDateS)
        If DateValue(xWDateS) > Now Then
            xlen = InStr(4, xWDateS, "/")
            xWDateS = DateAdd("yyyy", -1, xWDateS)
        End If
        txtFDateS = xWDateS
        xWDateS = DateAdd("yyyy", 1, xWDateS)
        xWDateS = DateAdd("d", -1, xWDateS)
        txtTDateS = xWDateS
    Else
        txtFDateS = ""
        txtTDateS = ""
        txtSICKT = 0
    End If
End If
If Not IsDate(txtfDate) And Not IsDate(txtFDateS) Then GoTo UpdByps

'If the following is not done esp. for Updated based on Entitlement Date (1) then it
'gives error (Type mismatch) in the following select statement
'Vacation & Sick date range from Entitlement Master since v7.6
'If Not IsDate(txtfDate) Then
'    txtfDate = glbCompEdFrom
'    txtTDate = glbCompEdTo
'End If
'
'If Not IsDate(txtFDateS) Then
'    txtFDateS = glbCompEdFromS
'    txtTDateS = glbCompEdToS
'End If

xVAC = 0
xSICK = 0

SQLQ = "SELECT SUM(AD_HRS) AS SUMHRS FROM HR_ATTENDANCE "
SQLQ = SQLQ & "WHERE AD_EMPNBR = " & lblEEID
SQLQ = SQLQ & "AND LEFT(AD_REASON,3)='VAC' "
If txtfDate.Text <> "" Then
    SQLQ = SQLQ & "AND AD_DOA>=" & Date_SQL(DateValue(txtfDate))
End If
If txtTDate.Text <> "" Then
    SQLQ = SQLQ & "AND AD_DOA<=" & Date_SQL(DateValue(txtTDate))
End If

rsAT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If IsNull(rsAT("SUMHRS")) Then xVAC = 0 Else xVAC = rsAT("SUMHRS")
rsAT.Close
SQLQ = "SELECT SUM(AD_HRS) AS SUMHRS FROM HR_ATTENDANCE "
SQLQ = SQLQ & "WHERE AD_EMPNBR = " & lblEEID
SQLQ = SQLQ & "AND LEFT(AD_REASON,3)='SIC' "
If txtFDateS.Text <> "" Then
    SQLQ = SQLQ & "AND AD_DOA>=" & Date_SQL(DateValue(txtFDateS))
End If
If txtTDateS.Text <> "" Then
    SQLQ = SQLQ & "AND AD_DOA<=" & Date_SQL(DateValue(txtTDateS))
End If

rsAT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If IsNull(rsAT("SUMHRS")) Then xSICK = 0 Else xSICK = rsAT("SUMHRS")
rsAT.Close

If glbCBrant Then
Dim rsEmpBack As New ADODB.Recordset
    SQLQ = "DELETE FROM HRVacBrant"
    gdbAdoIhr001B.BeginTrans
    gdbAdoIhr001B.Execute SQLQ
    gdbAdoIhr001B.CommitTrans
    rsEmpBack.Open "HRVacBrant", gdbAdoIhr001B, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
    rsEmpBack.AddNew
    rsEmpBack("ED_COMPNO") = "001"
    rsEmpBack("ED_EMPNBR") = lblEEID
    rsEmpBack("ED_PVAC") = Val(txtPVAC)
    rsEmpBack("ED_VAC") = Val(txtVAC)
    rsEmpBack("ED_VACT") = xtVACT
    If Not IsNull(xtEFDATE) And xtEFDATE <> "" Then
        rsEmpBack("ED_EFDATE") = CVDate(xtEFDATE)
    End If
    If Not IsNull(xtETDATE) And xtETDATE <> "" Then
        rsEmpBack("ED_ETDATE") = CVDate(xtETDATE)
    End If
    rsEmpBack("ED_REDATE") = Format(Now, "Short Date")
    rsEmpBack.Update
    rsEmpBack.Close
    
    txtPVAC = Val(txtPVAC) + Val(txtVAC) - xtVACT
    txtVAC = 0
End If
txtVACT = xVAC
txtSICKT = xSICK

UpdByps:
If glbEntOutStandingS$ = "2" Then lblDateS.Caption = dlpDate(7).Text
If glbEntOutStandingS$ = "3" Then lblDateS.Caption = dlpDate(6).Text
If glbEntOutStandingS$ = "4" Then lblDateS.Caption = dlpDate(5).Text
If glbEntOutStandingS$ = "5" Then lblDateS.Caption = dlpDate(3).Text
If glbEntOutStandingS$ = "6" Then lblDateS.Caption = dlpDate(4).Text

If glbEntOutStanding$ = "2" Then lblDATE.Caption = dlpDate(7).Text
If glbEntOutStanding$ = "3" Then lblDATE.Caption = dlpDate(6).Text
If glbEntOutStanding$ = "4" Then lblDATE.Caption = dlpDate(5).Text
If glbEntOutStanding$ = "5" Then lblDATE.Caption = dlpDate(3).Text
If glbEntOutStanding$ = "6" Then lblDATE.Caption = dlpDate(4).Text

glbENTScreen = True

End Function


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

Private Sub FMLA()

On Error GoTo Mod_Err
If glbCountry = "U.S.A." Then
    If gSec_Upd_Basic Then
        If glbWFC Then 'Ticket #6639
            If EmpCountry = "U.S.A." Then
                If IsDate(dlpDate(7).Text) And (Not IsDate(dlpDate(12).Text)) Then
                    dlpDate(12).Text = DateAdd("yyyy", 1, dlpDate(7).Text)
                End If
            End If
        Else
            If IsDate(dlpDate(7).Text) And (Not IsDate(dlpDate(12).Text)) Then
                dlpDate(12).Text = DateAdd("yyyy", 1, dlpDate(7).Text)
            End If
        End If
    End If
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREMP", "Modify")
Call RollBack '11Aug99 - js

End Sub

Private Function FldListOther()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ER_COMPNO,ER_EMPNBR, ER_PENSIONDATE1, ER_PENSIONDATE2, ER_PENSIONDATE3, ER_PENSIONDATE4,ER_PENSIONDATE5,ER_PENSIONDATE6, "
SQLQ = SQLQ & "ER_OTHERDATE1, ER_OTHERDATE2, ER_OTHERDATE3, ER_OTHERDATE4,ER_OTHERDATE5,ER_OTHERDATE6, "
SQLQ = SQLQ & "ER_OTHERDATE7, ER_OTHERDATE8, ER_OTHERDATE9, ER_OTHERDATE10, ER_COMMENT"
If glbtermopen Then
    SQLQ = SQLQ & ",ER_ID,TERM_SEQ"
End If
FldListOther = SQLQ
End Function


Private Function FldList()
Dim SQLQ

SQLQ = ""
SQLQ = SQLQ & "ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_PENSION, ED_EMP, "
SQLQ = SQLQ & "ED_EMPTYPE, ED_PT, ED_EFDATE, ED_EFDATES, ED_ORG,"
SQLQ = SQLQ & "ED_ETDATE, ED_ETDATES, ED_INTEL, ED_VACT, ED_VAC,ED_PVAC, ED_EMAIL,"
'SQLQ = SQLQ & "ED_LANG1, ED_SICKT, ED_LANG2, ED_DOH, ED_FDAY," 'George Apr 4,2006 #10574
SQLQ = SQLQ & "ED_SICKT, ED_DOH, ED_FDAY,"
SQLQ = SQLQ & "ED_ELIGIBLE, ED_SENDTE, ED_LDAY, ED_EARLYR,"
SQLQ = SQLQ & "ED_EXPYEAR, ED_OMERS, ED_NORMALR, ED_LTHIRE,"
SQLQ = SQLQ & "ED_USRDAT1, ED_LATESTR, ED_UNION, ED_FMLA, ED_DOB,"
SQLQ = SQLQ & "ED_SFDATE, ED_STDATE, ED_GLNO,"
SQLQ = SQLQ & "ED_DEPTEDATE,ED_DIVEDATE, ED_HIRECODE, ED_REGION, "  'Hemu Ticket 6549 - added ED_REGION
SQLQ = SQLQ & "ED_BENEFIT_GROUP, ED_LOC, ED_ADMINBY, ED_SECTION, "
SQLQ = SQLQ & "ED_PAYROLL_ID,"
SQLQ = SQLQ & "ED_UIC,ED_WCBCODE,"
SQLQ = SQLQ & "ED_SECTION,ED_SALDIST, ED_DEPTNO, ED_DOB, ED_DIV,"
If glbLinamar Then
    SQLQ = SQLQ & "ED_NOTELIGIBLE,"
End If
SQLQ = SQLQ & "ED_WITHSPOUSE , ED_LDATE, ED_LTIME, ED_LUSER"

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
SQLQ = SQLQ & ",ED_COUNTRY,ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1,ED_USER_NUM2,ED_PTEDATE,ED_SUPCODE,ED_ORGEDATE "

' Ticket #19266, wfc need these 2 fields, add them to all - Frank
''City of Kawartha Lakes
'If glbCompSerial = "S/N - 2363W" Then
'    SQLQ = SQLQ & ", ED_VADIM1" 'ED_PENSION,
'
''City of Timmins or City of Niagara Falls
'ElseIf glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
'    SQLQ = SQLQ & ",ED_VADIM2"
'End If
SQLQ = SQLQ & ",ED_VADIM1,ED_VADIM2"

'Ticket #22710 - County of Perth
'Ticket #29617 - Mississaugas of Scugog Island First Nation
If glbCompSerial = "S/N - 2417W" Or glbCompSerial = "S/N - 2485W" Then
    SQLQ = SQLQ & ",ED_VACPC"
End If

If glbCompSerial = "S/N - 2411W" Then  'Ticket #27899 - WDGPHU
    SQLQ = SQLQ & ",ED_ORGT1"
End If

FldList = SQLQ

End Function

Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer
'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If fglHredsem <> "" Then    'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'RFED'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And dlpDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "RFED"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew% = False And Edit1 = False And dlpDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "RFED"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
  
    If fglbNew% = False And Edit1 = True And dlpDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpDate(1).Text)
            dynHRAT("EF_FREAS") = "RFED"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglHredsem <> dlpDate(1).Text Then
            Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew% = False And Edit1 = True And dlpDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpDate(1).Text = "" Then
    updFollow = True
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


Private Function UPDStatusLOG()
On Error GoTo StatusLOGERR

Dim rsTB As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT TB_KEY FROM HRTABL "
SQLQ = SQLQ & " WHERE TB_NAME='EDEM' AND TB_USR3<>0 "
SQLQ = SQLQ & " AND TB_KEY IN ('" & SavEmp & "','" & clpCode(1).Text & "')"

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsTB.EOF Then GoTo MODNOUPD:
rsTB.Close

If SavEmp <> clpCode(1).Text Then GoTo MODUPD
If oFDate <> dlpDate(15).Text Then GoTo MODUPD
If OTDate <> dlpDate(16).Text Then GoTo MODUPD


GoTo MODNOUPD

MODUPD:

rsTB.Open "HRSTATUS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
rsTB.AddNew
rsTB("SC_COMPNO") = "001"
rsTB("SC_EMPNBR") = lblEEID
If IsDate(dlpDate(15).Text) Then rsTB("SC_FDATE") = dlpDate(15).Text
If IsDate(dlpDate(16).Text) Then rsTB("SC_TDATE") = dlpDate(16).Text
rsTB("SC_EMP_TABL") = "EDEM"
rsTB("SC_OLDEMP") = SavEmp
rsTB("SC_NEWEMP") = clpCode(1).Text
rsTB("SC_REASON_TABL") = "SCRE"
rsTB("SC_REASON") = "CHG"
rsTB("SC_LDATE") = Date
rsTB("SC_LTIME") = Time$
rsTB("SC_LUSER") = glbUserID
rsTB.Update


rsTB.Close

MODNOUPD:

Exit Function

StatusLOGERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING Status LOG RECORD", "Status LOG FILE", "UPDATE")
Call RollBack   '23June99 js

End Function

Private Sub dlpTermDate_LostFocus()
Call dlpDate_Change(6)
Call dlpDate_Change(7)
If glbWFC And glbtermopen Then 'Ticket #23948 Franks 06/24/2013
    If IsDate(dlpTermDate.Text) Then dlpDate(21).Text = dlpTermDate.Text
End If
End Sub
Public Sub Display_Value()
    Dim SQLQ
    
    If glbtermopen Then
       If rsDATA2.EOF Or rsDATA2.BOF Then Exit Sub
       Call Set_Control2("R", rsDATA2)
    End If
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    Call Set_Control("R", Me, rsDAT_Other, True)
    
    Call SET_UP_MODE
    
    Me.cmdModify_Click
    
    If glbCompSerial = "S/N - 2214W" Then 'Ticket #14550
        If clpPT.Text = "FT" Then
            lblIPhone.FontBold = True
        Else
            lblIPhone.FontBold = False
        End If
        locHOOPPBen = HOOPPBenFlag(glbLEE_ID)
        lblODate.FontBold = locHOOPPBen
    End If
    'Ticket #20136 - Franks 04/11/2011
    If glbWFC Then
        Call WFCDefaultEmailDisp 'Ticket #25248 Franks 03/24/2014
        Call chkWFCOtherDate6("Disp")
    End If
End Sub

Private Sub WFCDefaultEmailDisp() 'Ticket #25248 Franks 03/24/2014
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xSurname, XFNAME
Dim I As Integer

'"   For all countries, if Union = "NONE" or "EXEC", default the email address to be "first name_surname@woodbridgegroup.com". Ie: margaret_zyma@woodbridgegroup.com
'If NewHireForms.count > 0 Then
If Not glbtermopen Then
    If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
        If Len(txtEmail.Text) = 0 Then
            SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_FNAME,ED_ALIAS FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rs.EOF Then
                If IsNull(rs("ED_SURNAME")) Then xSurname = "" Else xSurname = rs("ED_SURNAME")
                
                'Ticket #29552 Franks 12/13/2016
                'If there are two surnames then use the first one
                xSurname = Trim(xSurname)
                
                I = InStr(1, xSurname, "-") 'check "-"
                If I > 0 Then
                    xSurname = Trim(Left(xSurname, I - 1))
                End If
                I = InStr(1, xSurname, " ") 'check space
                If I > 0 Then
                    xSurname = Trim(Left(xSurname, I))
                End If
                
                'If Alias is populate use it as Frist Name
                XFNAME = ""
                If Not IsNull(rs("ED_ALIAS")) Then 'xFName = "" Else xFName = rs("ED_FNAME")
                    If Len(Trim(rs("ED_ALIAS"))) > 0 Then
                        XFNAME = Trim(rs("ED_ALIAS"))
                    End If
                End If
                If Len(XFNAME) = 0 Then
                    If IsNull(rs("ED_FNAME")) Then XFNAME = "" Else XFNAME = rs("ED_FNAME")
                End If
                
                'If there are two first names then use the first one
                XFNAME = Trim(XFNAME)
                I = InStr(1, XFNAME, "-") 'check "-"
                If I > 0 Then
                    XFNAME = Trim(Left(XFNAME, I - 1))
                End If
                I = InStr(1, XFNAME, " ") 'check space
                If I > 0 Then
                    XFNAME = Trim(Left(XFNAME, I))
                End If
                'Ticket #29552 Franks 12/13/2016 - end
                
                txtEmail.Text = Trim(XFNAME) & "_" & Trim(xSurname) & "@woodbridgegroup.com"
            End If
        End If
    End If
End If
'End If
End Sub

Private Sub chkWFCOtherDate6(xlType)
lblWFCMsg.Visible = False
If glbEmpCountry = "U.S.A." Then
    If Len(clpVadim2.Text) > 0 Then 'Pay Group
        'If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
        'Ticket #25248 Franks 03/24/2014 - both Salary and Hourly
        'Ticket #25352 Franks 04/16/2014 - exclude "COOP" and "STUD"
        If Not (clpCode(1).Text = "COOP" Or clpCode(1).Text = "STUD" Or clpCode(1).Text = "CONP") Then      'Ticket #29660 - Ignore CONP
            If xlType = "Disp" Then
            'If Other Date 6 is blank, display a message in RED under NGS Sub Group saying “No 401k Eligibility Date
                If Len(dlpDate(29).Text) = 0 Then
                    'lblWFCMsg.Caption = "Enter 401k Eligibility Date on the Other Dates Tab"
                    'Ticket #25248 Franks 03/24/2014
                    lblWFCMsg.Caption = "401k Eligibility Date is missing - Enter on the Other Dates Tab and Save"
                    lblWFCMsg.Top = lblVadim11.Top + 610 '
                    lblWFCMsg.Left = 4500 'lblVadim11.Left
                    lblWFCMsg.Visible = True
                Else
                    lblWFCMsg.Caption = ""
                    lblWFCMsg.Visible = False
                End If
            End If
            If xlType = "Upt" Then
                If Len(dlpDate(29).Text) = 0 Then
                    dlpDate(29).Text = dlpDate(7).Text 'DOH
                End If
            End If
        End If
        'End If
    End If
End If
End Sub

Private Sub Set_Control2(Act As String, Optional rsTA As ADODB.Recordset)

If Act = "U" Then
    If Len(dlpTermDate.Text) = 0 Then
        rsTA!term_dot = Null
    Else
        rsTA!term_dot = dlpTermDate.Text
    End If
    
    If Len(clpCode(5).Text) = 0 Then
        rsTA!term_reason = Null
    Else
        rsTA!term_reason = clpCode(5).Text
    End If
    
    If Len(dlpRehired.Text) = 0 Then
        rsTA!Term_DOR = Null
    Else
        rsTA!Term_DOR = dlpRehired.Text
    End If
    
    If Len(txtComments.Text) = 0 Then
       rsTA("Term_Comments") = Null
    Else
       rsTA("Term_Comments") = txtComments.Text
    End If
    
    rsTA!Term_Rehire = chkRehire
    
    If glbWFC Then 'Ticket #15248
        If Len(clpCode(3).Text) = 0 Then
            rsTA!Term_Cause = Null
        Else
            rsTA!Term_Cause = clpCode(3).Text
        End If
    End If
ElseIf Act = "B" Then
    dlpTermDate.Text = ""
    clpCode(5).Text = ""
    dlpRehired.Text = ""
    txtComments.Text = ""
    chkRehire = False
ElseIf Act = "R" Then
    dlpTermDate.Text = ""
    clpCode(5).Text = ""
    dlpRehired.Text = ""
    txtComments.Text = ""
    chkRehire = False
    If glbWFC Then 'Ticket #15248
        clpCode(3).Text = ""
    End If
    If rsTA.EOF Or rsTA.BOF Then Exit Sub
        If IsNull(rsTA!term_dot) Then
            dlpTermDate.Text = ""
        Else
            dlpTermDate.Text = rsTA!term_dot
        End If
        If IsNull(rsTA!term_reason) Then
            clpCode(5).Text = ""
        Else
            clpCode(5).Text = rsTA!term_reason
        End If
        If IsNull(rsTA!Term_DOR) Then
            dlpRehired.Text = ""
        Else
            dlpRehired.Text = rsTA!Term_DOR
        End If
        If IsNull(rsTA("Term_Comments")) Then
             txtComments.Text = ""
        Else
             txtComments.Text = rsTA("Term_Comments")
        End If
        If Not IsNull(rsTA!Term_Rehire) Then
            chkRehire = rsTA!Term_Rehire
        End If
        If glbWFC Then 'Ticket #15248
            If IsNull(rsTA!Term_Cause) Then
                clpCode(3).Text = ""
            Else
                clpCode(3).Text = rsTA!Term_Cause
            End If
        End If
        'Ticket #24317 Franks 09/18/2013
        lblUserDesc.Caption = GetUserDesc(rsTA!Term_LUSER)
        dlpDate(35).Text = rsTA!Term_LDATE
    End If
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
'If glbtermopen Then
'    RelateMode = RelateTermEmp
'Else
    RelateMode = RelateEMP
'End If
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Basic
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
Call modSTUPD(TF)
''If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/20/2014
''    TF = getFamilyDayUpdateRight(UpdateRight, glbLEE_ID)
''Else
''    If Not UpdateRight Then TF = False
''End If
'Ticket #24729 Frank 01/28/2014

If Not UpdateRight Then TF = False

Call modSTUPD(TF)

'George on Jan 26,2006 #10266
glbDocName = "Resume" 'George on Jan 26,2006 #10266
If gsAttachment_DB Then 'George on Jan 24,2006 #10266
    Call DispimgIcon(Me, "frmEESTATS")
End If
If glbtermopen Then
    glbDocName = "Termination" 'George on Jan 26,2006 #10266
    If gsAttachment_DB Then 'George on Jan 24,2006 #10266
        Call DispimgIcon(Me, "frmEESTATS")
    End If
End If

'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        If glbtermopen Then
            cmdImport1.Visible = False
            cmdLOAComments.Visible = False
        Else
            cmdImport.Visible = False
                        
            'Release 8.1
            cmdImport2.Visible = False
            cmdLOAComments.Visible = False
        End If
    Else
        If glbtermopen Then
            cmdImport1.Visible = True
            cmdLOAComments.Visible = False
        Else
            cmdImport.Visible = True
            lblImport1.Visible = False
            imgNoSec1.Visible = False
            imgSec1.Visible = False
            
            'Release 8.1
            'Show/Hide the LOA Attachment buttons depending on the Empployment Status Code
            Call ShowHide_LOA_Attachment_Buttons
        End If
    End If
    If Not (gSec_Upd_Basic And Not glbtermopen) Then
        If glbtermopen Then
            cmdImport1.Visible = False
        Else
            cmdImport.Visible = False
        End If
    End If
End If
'George on Jan 26,2006 #10266
'added by Bryan Nov 10, 2006, Ticket #12065
If Not glbtermopen Then
    lblTitle(16).Visible = False
    lblTitle(17).Visible = False
    lblTitle(18).Visible = False
    lblTitle(19).Visible = False
    lblRehired(1).Visible = False
    lblTitle(2).Visible = False
    dlpTermDate.Visible = False
    dlpRehired.Visible = False
    chkRehire.Visible = False
    clpCode(5).Visible = False
    clpCode(3).Visible = False 'for wfc
    txtComments.Visible = False
    'Ticket #24317 Franks 09/18/2013 - begin
    lblUpdateBy.Visible = False
    lblUserDesc.Visible = False
    lblTitle(3).Visible = False
    dlpDate(35).Visible = False
    'Ticket #24317 Franks 09/18/2013 - end
Else
    lblTitle(16).Visible = True
    lblTitle(17).Visible = True
    lblTitle(18).Visible = True
    lblTitle(19).Visible = True
    lblRehired(1).Visible = True
    dlpTermDate.Visible = True
    dlpRehired.Visible = True
    chkRehire.Visible = True
    clpCode(5).Visible = True
    txtComments.Visible = True
    'Ticket #24317 Franks 09/18/2013 - begin
    lblUpdateBy.Visible = True
    lblUserDesc.Visible = True
    lblTitle(3).Visible = True
    dlpDate(35).Visible = True
    'Ticket #24317 Franks 09/18/2013 - end
End If

End Sub



Private Function VadimControl(Action)
Dim rsVadim As New ADODB.Recordset
Dim ctlName As Control
Dim lblName As Label
VadimControl = False
If Action = "Check" Then
    Data1.Recordset.Requery
    If Data1.Recordset(Vadim_PayType_field) = "S" Then
        If Len(clpSalDist) = 0 And glbCompSerial <> "S/N - 2362W" Then
            MsgBox lStr(lblSalDist & " is a required field for salaried employee.")
            clpSalDist.SetFocus
            Exit Function
        End If
    Else
        If Len(clpSalDist) > 0 Then
            If InStr("H,C,P,F,", Data1.Recordset(Vadim_PayType_field) & ",") Then
                MsgBox lStr(lblSalDist & " must be empty for hourly employee.")
                clpSalDist.SetFocus
                Exit Function
            End If
        End If
    End If
End If
rsVadim.Open "SELECT * FROM VADIM_MAPPING WHERE VADIM_FIELD IN ('Start Date')", gdbAdoIhr001, adOpenForwardOnly
Do Until rsVadim.EOF
    Select Case rsVadim("INFOHR_FIELD")
    Case "ED_SENDTE"
        Set ctlName = dlpDate(6)
        Set lblName = lblSen
    Case "ED_LTHIRE"
        Set ctlName = dlpDate(5)
        Set lblName = lblLHire
    Case "ED_UNION"
        Set ctlName = dlpDate(4)
        Set lblName = lblUDate
    Case "ED_FDAY"
        Set ctlName = dlpDate(0)
        Set lblName = lblFDay
    Case "ED_LDAY"
        Set ctlName = dlpDate(1)
        Set lblName = lblLDay
    Case "ED_USRDAT1"
        Set ctlName = dlpDate(3)
        Set lblName = lblUDay
    End Select
    If lblName Is Nothing Or ctlName Is Nothing Then VadimControl = True: Exit Function

    If Not rsVadim("INFOHR_FIELD") = "ED_DOH" Then
        If Action = "Show" Then
            lblName.FontBold = True
        ElseIf Action = "Check" Then
            If Len(ctlName) = 0 Then
                MsgBox lStr(lblName & " is a required field.")
                ctlName.SetFocus
                Exit Function
            End If
        End If
    End If
    rsVadim.MoveNext
Loop
VadimControl = True
End Function

Private Sub updFollowUserDefinedDate() 'WFC only
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset

On Error GoTo CrFollow_Err

If ODate(3) = dlpDate(3) Then Exit Sub

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
If IsDate(ODate(3)) Then
    SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS='WAED' "
    SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(ODate(3))
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not IsDate(ODate(3)) Or rsTB.EOF Then
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
Else
    If Not IsDate(dlpDate(3).Text) Then
        rsTB.Delete
        Msg = "A Follow Up Record was deleted!"
        GoTo TheEnd
    Else
        Msg = "A Follow Up Record was updated!"
    End If
End If
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = glbLEE_ID
rsTB("EF_FDATE") = CVDate(dlpDate(3).Text)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
End If
rsTB("EF_FREAS") = "WAED"
rsTB("EF_COMMENTS") = "WORK AUTHORIZATION EXPIRE DATE"
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update

TheEnd:
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next
End Sub

Private Function WSQLQ()
'FIXED BY SAM AS IT WAS READING WRONG TABLE 08/16/06
If Len(rsTA("ED_DEPTNO")) > 0 Then WSQLQ = WSQLQ & " WHERE EL_DEPTNO = '" & rsTA("ED_DEPTNO") & "'"
If Len(rsTA("ED_DIV")) > 0 Then WSQLQ = WSQLQ & " AND EL_DIV = '" & rsTA("ED_DIV") & "' "
If Len(rsTA("ED_LOC")) > 0 Then WSQLQ = WSQLQ & " AND EL_LOC = '" & rsTA("ED_LOC") & "' "
If Len(rsTA("ED_ORG")) > 0 Then WSQLQ = WSQLQ & " AND EL_ORG = '" & rsTA("ED_ORG") & "' "
If (rsTA("ED_EML")) > 0 Then WSQLQ = WSQLQ & " AND EL_EMP = '" & rsTA("ED_EML") & "' "
If (rsTA("ED_PT")) > 0 Then WSQLQ = WSQLQ & " AND EL_PT = '" & rsTA("ED_PT") & "' "

End Function
Private Function WSQLQ2()
'FIXED BY SAM AS IT WAS READING WRONG TABLE 08/16/06

WSQLQ2 = WSQLQ2 & " WHERE EL_DEPTNO = ''"
WSQLQ2 = WSQLQ2 & " AND EL_DIV =''"
WSQLQ2 = WSQLQ2 & " AND EL_LOC =''"
WSQLQ2 = WSQLQ2 & " AND EL_ORG =''"
WSQLQ2 = WSQLQ2 & " AND EL_EMP =''"
WSQLQ2 = WSQLQ2 & " AND EL_PT =''"

End Function

Private Function WSQLQ1()
'COMMENTED BY SAM AS IT DOES NOT MAKE SENSE  08/16/06
'WSQLQ1 = WSQLQ1 & " WHERE " & glbSeleDeptUn
If Len(rsTA("ED_DEPTNO")) > 0 Then WSQLQ1 = WSQLQ1 & " WHERE ED_DEPTNO = '" & rsTA("ED_DEPTNO") & "'"
If Len(rsTA("ED_DIV")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_DIV = '" & rsTA("ED_DIV") & "' "
If Len(rsTA("ED_LOC")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_LOC = '" & rsTA("ED_LOC") & "' "
If Len(rsTA("ED_ORG")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_ORG = '" & rsTA("ED_ORG") & "' "
If (rsTA("ED_EML")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_EMP = '" & rsTA("ED_EML") & "' "
If (rsTA("ED_PT")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_PT = '" & rsTA("ED_PT") & "' "

End Function

Private Function UPDEML() 'Ticket #13412 Frank Jul 24, 2007
Dim rsEmpEml As New ADODB.Recordset
Dim rsEML As New ADODB.Recordset
Dim SQLQ
Dim xEmlVal As Double

On Error GoTo ErrorHandler

    xEmlVal = 0
    SQLQ = "SELECT * FROM HR_EMLSETUP "
    rsEML.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsEML.EOF
        SQLQ = "SELECT ED_EMPNBR, ED_EML FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID & " "
        If Not IsNull(rsEML("EL_DIV")) Then
            If Len(rsEML("EL_DIV")) > 0 Then
                SQLQ = SQLQ & "AND ED_DIV = '" & rsEML("EL_DIV") & "' "
            End If
        End If
        If Not IsNull(rsEML("EL_DEPTNO")) Then
            If Len(rsEML("EL_DEPTNO")) > 0 Then
                SQLQ = SQLQ & "AND ED_DEPTNO = '" & rsEML("EL_DEPTNO") & "' "
            End If
        End If
        If Not IsNull(rsEML("EL_ORG")) Then
            If Len(rsEML("EL_ORG")) > 0 Then
                SQLQ = SQLQ & "AND ED_ORG = '" & rsEML("EL_ORG") & "' "
            End If
        End If
        If Not IsNull(rsEML("EL_LOC")) Then
            If Len(rsEML("EL_LOC")) > 0 Then
                SQLQ = SQLQ & "AND ED_LOC = '" & rsEML("EL_LOC") & "' "
            End If
        End If
        If Not IsNull(rsEML("EL_EMP")) Then
            If Len(rsEML("EL_EMP")) > 0 Then
                SQLQ = SQLQ & "AND ED_EMP = '" & rsEML("EL_EMP") & "' "
            End If
        End If
        If Not IsNull(rsEML("EL_PT")) Then
            If Len(rsEML("EL_PT")) > 0 Then
                SQLQ = SQLQ & "AND ED_PT = '" & rsEML("EL_PT") & "' "
            End If
        End If
        rsEmpEml.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpEml.EOF Then
            xEmlVal = rsEML("EL_EML")
            rsEmpEml("ED_EML") = xEmlVal
            rsEmpEml.Update
        End If
        rsEmpEml.Close
        
        rsEML.MoveNext
    Loop
    rsEML.Close
    If xEmlVal = 0 Then
        SQLQ = "UPDATE HREMP SET ED_EML = " & xEmlVal & " WHERE ED_EMPNBR =" & glbLEE_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    
    Exit Function
ErrorHandler:
glbFrmCaption$ = "EML Entitlement on new hire"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UPDEML", "", "EML")
If gintRollBack% = False Then
    Resume Next
End If

End Function

Private Sub UPDOvertime_Overview()
Dim rsHREmp As New ADODB.Recordset
Dim rsOvtMst As New ADODB.Recordset
Dim rsOvtEmp As New ADODB.Recordset
Dim SQLQ As String
Dim flgUpdated As Boolean

On Error GoTo Err_UPDOvertime_Overview

    'Ticket #22847- To see if employee record updated, and if employee needs to be in the Overtime Bank master based on the rule
    flgUpdated = False

    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER ORDER BY OM_ORG"
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsOvtMst.EOF
        Set rsHREmp = Nothing
        SQLQ = "SELECT ED_EMPNBR, ED_EMP,ED_PT,ED_ORG,ED_LOC,ED_REGION,ED_ADMINBY,ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID & " "
        'SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHREmp.EOF Then
            'Ticket #15753 - from OM_LOC to OM_SECTION
            If ((IsNull(rsOvtMst("OM_EMP")) Or rsOvtMst("OM_EMP") = "") Or (rsHREmp("ED_EMP") = rsOvtMst("OM_EMP"))) And _
                ((IsNull(rsOvtMst("OM_PT")) Or rsOvtMst("OM_PT") = "") Or (rsHREmp("ED_PT") = rsOvtMst("OM_PT"))) And _
                ((IsNull(rsOvtMst("OM_LOC")) Or rsOvtMst("OM_LOC") = "") Or (rsHREmp("ED_LOC") = rsOvtMst("OM_LOC"))) And _
                ((IsNull(rsOvtMst("OM_REGION")) Or rsOvtMst("OM_REGION") = "") Or (rsHREmp("ED_REGION") = rsOvtMst("OM_REGION"))) And _
                ((IsNull(rsOvtMst("OM_ADMINBY")) Or rsOvtMst("OM_ADMINBY") = "") Or (rsHREmp("ED_ADMINBY") = rsOvtMst("OM_ADMINBY"))) And _
                ((IsNull(rsOvtMst("OM_SECTION")) Or rsOvtMst("OM_SECTION") = "") Or (rsHREmp("ED_SECTION") = rsOvtMst("OM_SECTION"))) And _
                ((IsNull(rsOvtMst("OM_ORG")) Or rsOvtMst("OM_ORG") = "") Or (rsHREmp("ED_ORG") = rsOvtMst("OM_ORG"))) Then
                
                    Set rsOvtEmp = Nothing
                    rsOvtEmp.Open "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR=" & rsHREmp("ED_EMPNBR"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsOvtEmp.EOF Then
                        rsOvtEmp.AddNew
                        rsOvtEmp("OT_PBANK") = 0
                    End If
                    rsOvtEmp("OT_COMPNO") = "001"
                    rsOvtEmp("OT_EMPNBR") = rsHREmp("ED_EMPNBR")
                    rsOvtEmp("OT_BANK") = Get_OvertimeBank(rsHREmp("ED_EMPNBR"), rsOvtMst("OM_EFDATE"), rsOvtMst("OM_ETDATE")) * Val(rsOvtMst("OM_MULTIPLIER"))
                    rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(rsHREmp("ED_EMPNBR"), rsOvtMst("OM_EFDATE"), rsOvtMst("OM_ETDATE"))
                    rsOvtEmp("OT_MBANK") = rsOvtMst("OM_MAX_BANK_HRS")
                    rsOvtEmp("OT_EFDATE") = rsOvtMst("OM_EFDATE")   'Format("1/1/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_ETDATE") = rsOvtMst("OM_ETDATE")   'Format("12/31/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_LDATE") = Date
                    rsOvtEmp("OT_LTIME") = Time$
                    rsOvtEmp("OT_LUSER") = glbUserID
                    rsOvtEmp.Update
                    
                    rsOvtEmp.Close
                    
                    'Ticket #22847- Employee Overtime Bank record updated - employee part of the rule.
                    flgUpdated = True
                    
                    Exit Do
            End If
        End If
        rsHREmp.Close
        
        rsOvtMst.MoveNext
    Loop
    rsOvtMst.Close
    
    'Ticket #22847 - Check if Employee was part of the Overtime Bank Master rule
    If flgUpdated = False Then
        'Employee was not part of the Overtime Bank Master rule.
        'Delete the Overtime Bank record if it exists for this employee.
        gdbAdoIhr001.Execute "DELETE HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
    End If
    
    Exit Sub

Err_UPDOvertime_Overview:
glbFrmCaption$ = "Overtime Bank Overview update on new hire"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UPUPDOvertime_Overview", "", "EML")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Private Sub updFollowPension() 'For Walter Fedy
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err

If Not IsDate(dlpDate(7).Text) Then Exit Sub

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
SQLQ = SQLQ & " WHERE EF_EMPNBR=" & glbLEE_ID
SQLQ = SQLQ & " AND EF_FREAS='PENS' "
SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(dlpDate(7).Text)

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsTB.EOF Then
    Exit Sub
End If
rsTB.AddNew
Msg = "A Follow Up Record was created!"

rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = glbLEE_ID
rsTB("EF_FDATE") = CVDate(DateAdd("YYYY", 1, dlpDate(7).Text))
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
End If
rsTB("EF_FREAS") = "PENS"
rsTB("EF_COMMENTS") = ""
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update


Dim rsTT As New ADODB.Recordset
rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='PENS'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsTT.EOF Then
    rsTT.AddNew
    rsTT("TB_COMPNO") = "001"
    rsTT("TB_NAME") = "FURE"
    rsTT("TB_KEY") = "PENS"
    rsTT("TB_DESC") = "Pension Followup"
    rsTT("TB_LUSER") = glbUserID
    rsTT("TB_LDATE") = Date
    rsTT("TB_LTIME") = Time$
    rsTT.Update
End If
rsTT.Close

'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
'follow up record
Call Grant_FollowUpCode_Security(glbUserID, "PENS", "Pension Followup")

TheEnd:
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next
End Sub

Private Sub updFollowSin()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err
If ODate(3) = dlpDate(3) Then Exit Sub
SQLQ = "SELECT * FROM HR_FOLLOW_UP "
If IsDate(ODate(3)) Then
    SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS='SIN' "
    SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(ODate(3))
End If

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not IsDate(ODate(3)) Or rsTB.EOF Then
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
Else
    If Not IsDate(dlpDate(3).Text) Then
        rsTB.Delete
        Msg = "A Follow Up Record was deleted!"
        GoTo TheEnd
    Else
        Msg = "A Follow Up Record was updated!"
    End If
End If
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = glbLEE_ID
rsTB("EF_FDATE") = CVDate(dlpDate(3).Text)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
End If
rsTB("EF_FREAS") = "SIN"
rsTB("EF_COMMENTS") = ""
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update


Dim rsTT As New ADODB.Recordset
rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='SIN'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsTT.EOF Then
    rsTT.AddNew
    rsTT("TB_COMPNO") = "001"
    rsTT("TB_NAME") = "FURE"
    rsTT("TB_KEY") = "SIN"
    rsTT("TB_DESC") = "SIN# Expiration"
    rsTT("TB_LUSER") = glbUserID
    rsTT("TB_LDATE") = Date
    rsTT("TB_LTIME") = Time$
    rsTT.Update
End If
rsTT.Close

'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
'follow up record
Call Grant_FollowUpCode_Security(glbUserID, "SIN", "SIN# Expiration")

TheEnd:
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub
'
'
'Public Function GetShortName(ByVal sLongFileName As String) As String
'    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
'    'Set up buffer area for API function call return
'    sShortPathName = Space(255)
'    iLen = Len(sShortPathName)
'
'    'Call the function
'    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
'    'Strip away unwanted characters.
'    GetShortName = Left(sShortPathName, lRetVal)
'End Function

Private Sub Pass_TermEmp_Change_Vadim()
Dim UpdateAudit As Boolean

UpdateAudit = False
Dim HRChanges As New Collection

If isChanged_Field(HRChanges, SavEmp, clpCode(1)) Then UpdateAudit = True
If isChanged_Field(HRChanges, OEmptype, txtEmpType) Then UpdateAudit = True
If isChanged_Field(HRChanges, SavPT, clpPT) Then UpdateAudit = True

'Town of Aurora - Do not transfer Union Code for Non Union.
If glbCompSerial = "S/N - 2378W" Then
    If clpCode(2).Text <> "0" Then
        If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
    ElseIf SavOrg <> "" And SavOrg <> "0" Then
        If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
    End If
Else
    If isChanged_Field(HRChanges, SavOrg, clpCode(2)) Then UpdateAudit = True
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If SavOrg <> clpCode(2) Then
        If isChanged_Field(HRChanges, "", txtVadim1) Then UpdateAudit = True
    End If
End If

If isChanged_Field(HRChanges, ODate(8), dlpDate(0)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(1), dlpDate(1)) Then UpdateAudit = True

'Ticket #25469 - City of Campbell River - No logic behind OMERS date
If glbCompSerial <> "S/N - 2458W" Then
    If isChanged_Field(HRChanges, ODate(2), dlpDate(2)) Then UpdateAudit = True
End If

'City of Kawartha Lakes - OMERS Date and RPP# logic
If glbCompSerial = "S/N - 2363W" Then
    If ODate(2) <> dlpDate(2) Then
        If isChanged_Field(HRChanges, "", txtRPP) Then UpdateAudit = True
    End If
End If

'City of Timmins   - RPP# logic
If glbCompSerial = "S/N - 2375W" Then
    If isChanged_Field(HRChanges, "", txtRPP) Then UpdateAudit = True
End If

'Town of Aurora - Ticket #20931 - as per mapping documentation
'City of Niagara Falls - Ticket #20053 - Transfer Benefit Group Code to EMP_CLASS_CODE
If glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2378W" Then
    If isChanged_Field(HRChanges, OBenGrp, clpBGroup) Then UpdateAudit = True
End If

'Ticket #24996 - City of Campbell River - Transfer ED_SECTION to EMP_CLASS_CODE and Benefit Group to EMP_DEFAULT_JOB
If glbCompSerial = "S/N - 2458W" Then
    If isChanged_Field(HRChanges, OBenGrp, clpBGroup) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OSection, clpCode(4)) Then UpdateAudit = True
    
    'For Sick and Vacation Accruals
    If isChanged_Field(HRChanges, oUSER_NUM1, txtUserNum1) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oUSER_NUM2, txtUserNum2) Then UpdateAudit = True
End If

If isChanged_Field(HRChanges, ODate(4), dlpDate(4)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(5), dlpDate(5)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(6), dlpDate(6)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(7), dlpDate(7)) Then UpdateAudit = True

If glbCompSerial <> "S/N - 2375W" Then   'City of Timmins
    If isChanged_Field(HRChanges, ODate(3), dlpDate(3)) Then UpdateAudit = True
End If

If isChanged_Field(HRChanges, ODate(13), dlpDate(12)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(9), dlpDate(8)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(10), dlpDate(9)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(11), dlpDate(10)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODate(12), dlpDate(11)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODeptEDate, dlpDate(13)) Then UpdateAudit = True
If isChanged_Field(HRChanges, ODivEdate, dlpDate(14)) Then UpdateAudit = True
If isChanged_Field(HRChanges, oFDate, dlpDate(15)) Then UpdateAudit = True
If isChanged_Field(HRChanges, OTDate, dlpDate(16)) Then UpdateAudit = True
If isChanged_Field(HRChanges, OINTEL, txtIPHONE) Then UpdateAudit = True
'If isChanged_Field(HRChanges, OLANG1, clpCode(3)) Then UpdateAudit = True
'If isChanged_Field(HRChanges, OLANG2, clpCode(4)) Then UpdateAudit = True
If isChanged_Field(HRChanges, oHireCode, clpCode(6)) Then UpdateAudit = True
If isChanged_Field(HRChanges, oSalDist, clpSalDist) Then UpdateAudit = True
If isChanged_Field(HRChanges, oEmail, txtEmail) Then UpdateAudit = True
If isChanged_Field(HRChanges, OWITHSPOUSE, chkSpouse) Then UpdateAudit = True
If isChanged_Field(HRChanges, OEXPYEAR, txtExpYear) Then UpdateAudit = True

Call Passing_Changes(HRChanges, Status, "M", Date, glbTERM_ID, rsDATA("ED_PAYROLL_ID"))

End Sub

Private Sub txtUserNum1_Change()
If glbWFC Then 'Ticket #22448 Franks
    Call PopcomUserText2(txtUserNum1.Text)
End If
End Sub

Private Sub txtUserNum1_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserNum2_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserText1_Change()
'Ticket #23537 - Essex Country Lib. - Remove this logic now
'Ticket #18789 Franks 05/06/2011
'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
'    comUserText1.ListIndex = -1
'    comUserText1.ListIndex = GetUserText1Index(txtUserText1.Text)
'End If
'Ticket #24976 - VitalAire Canada Inc.
If glbCompSerial = "S/N - 2380W" Then
    comUserText1.ListIndex = -1
    comUserText1.ListIndex = GetUserText1Index(Trim(txtUserText1.Text))
End If
End Sub

Private Sub txtUserText1_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Sub ADP_Control(xEmpNo, xCtl1_Val)
Dim rsADP As New ADODB.Recordset
Dim rsADP_AUDIT As New ADODB.Recordset
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim xFLAG As Boolean
Dim strFields As String
Dim x As Integer

    SQLQ = "Select * from HR_ADP "
    SQLQ = SQLQ & " where AP_EMPNBR = " & xEmpNo
    rsADP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsADP.EOF Then
        rsADP.AddNew
        rsADP("AP_COMPNO") = "001"
        rsADP("AP_EMPNBR") = xEmpNo
        rsADP("AP_DCP1") = "0"
        rsADP("AP_LUSER") = glbUserID
        rsADP("AP_LDATE") = Date
        rsADP("AP_LTIME") = Time$
        rsADP.Update
    End If
    rsADP.Close
    
    'Audit
    rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    
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
    'strfields added by Bryan 02/Dec/05 Ticket#9899
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
    strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, "
    strFields = strFields & "AU_ADP_FLAG, "
    strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE "
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    xADD = False
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
    rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
    rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    
    rsTA("AU_ADP_FLAG") = True
    
   
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M" 'ACTX
    rsTA.Update
    rsTA.Close
End Sub
Private Function AUDIT_NGS_TRANS()
Dim rsTA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim SQLQ As String
Dim xBenTermDate
Dim xEmpStatCode
Dim xUptFlag As Boolean
Dim xDate1, xDate2
Dim xLDate
Dim xForm As String
Dim xEmpID
Dim xUptType 'Ticket #20385 Franks 05/31/2011
Dim xExitFlag As Boolean
Dim xOldNGSVal, xNewNGSVal 'Ticket #22718

On Error GoTo AUDIT_ERR
If Not glbNGS_OnFlag Then
    AUDIT_NGS_TRANS = True
    Exit Function
End If

AUDIT_NGS_TRANS = False

glbEmpDiv = rsDATA("ED_DIV")
glbUNION = clpCode(2).Text
glbWFCNGSSubGroup = clpVadim1.Text
glbWFCPayGroup = clpVadim2.Text

xLDate = Date

If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
    xEmpID = glbTERM_ID
Else
    xEmpID = glbLEE_ID
End If

If NewHireForms.count > 0 Then
    xUptType = "A"
Else
    xUptType = "M"
End If

''No NGS Sub Group, skip
'If Len(clpVadim1.Text) = 0 Then Exit Function
''Ticket #22663 Franks 10/16/2012 'ED_PT is not FT, skip
'If Not clpPT.Text = "FT" Then Exit Function

xExitFlag = False
'No NGS Sub Group, skip
If Len(clpVadim1.Text) = 0 Then xExitFlag = True
''Ticket #22663 Franks 10/16/2012 'ED_PT is not FT, skip
'If Not clpPT.Text = "FT" Then xExitFlag = True
'Ticket #22991 Franks 01/03/2013 - add PT as part of NGS
If Not (clpPT.Text = "FT" Or clpPT.Text = "PT") Then xExitFlag = True

'Ticket #22699 if NGS End Date entered then create a NGS audit record, then exit function
'NGS End Date change - begin
If Len(OVadim11) > 0 Then
    If IsDate(oOTHERDATE2) Then xDate1 = CVDate(oOTHERDATE2) Else xDate1 = ""
    If IsDate(dlpDate(25).Text) Then xDate2 = CVDate(dlpDate(25).Text) Else xDate2 = ""
    If Not (xDate1 = xDate2) Then
        If IsDate(xDate2) Then
            If CVDate(xDate2) > CVDate(Date) Then xLDate = CVDate(xDate2)
        End If
        Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Other Date 2"), xDate1, xDate2, xLDate)
        Exit Function
    End If
End If
'NGS End Date change - end

If xExitFlag Then
    '''Ticket #22699 if NGS End Date entered then create a NGS audit record, then exit function
    '''NGS End Date change - begin
    ''If Len(OVadim11) > 0 Then
    ''    If IsDate(oOTHERDATE2) Then xDate1 = CVDate(oOTHERDATE2) Else xDate1 = ""
    ''    If IsDate(dlpDate(25).Text) Then xDate2 = CVDate(dlpDate(25).Text) Else xDate2 = ""
    ''    If Not (xDate1 = xDate2) Then
    ''        If IsDate(xDate2) Then
    ''            If CVDate(xDate2) > CVDate(Date) Then xLDate = CVDate(xDate2)
    ''        End If
    ''        Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Other Date 2"), xDate1, xDate2, xLDate)
    ''    End If
    ''End If
    '''NGS End Date change - end
    Exit Function
End If

'Ticket #20385 Franks 05/31/2011
'Change IHR to write to the NGS Audit Table if the employee has a NGS Sub Group
'regardless of entering a Start Date.
'If IsDate(dlpDate(24).Text) Then
'    If CVDate((dlpDate(24).Text)) > CVDate(Date) Then
'        xLDate = CVDate((dlpDate(24).Text))
'    End If
'End If
    
''New Hire only
'If NewHireForms.count > 0 Then
'    'call one 'public function to add these fields
'    Call AUDIT_NGS_NEWHIRE(glbLEE_ID, xLDate)
'    AUDIT_NGS_TRANS = True
'    Exit Function
'End If

'Ticket #20385 Franks 05/31/2011
''No NGS Effective Date, skip
'If Len(dlpDate(24).Text) = 0 Then Exit Function

'NGS field changes --------------------------------------
    
'NGS Start Date change - begin
If IsDate(oOTHERDATE1) Then
    xDate1 = CVDate(oOTHERDATE1)
Else
    xDate1 = ""
End If
If IsDate(dlpDate(24).Text) Then
    xDate2 = CVDate(dlpDate(24).Text)
Else
    xDate2 = ""
End If
If Not (xDate1 = xDate2) Then
    xForm = "Status/Dates"
    'Ticket #20385 Franks 05/31/2011
    ''If Len(xDate1) = 0 Then
    ''    xForm = "New Hire"
    ''End If
    ''If NewHireForms.count > 0 Then
    ''    Call NGSAuditAdd(xEmpID, xUptType, xForm, lStr("Other Date 1"), xDate1, xDate2, xLDate)
    ''    AUDIT_NGS_TRANS = True
    ''    Exit Function 'new hire only need this record for trigger NGS export
    ''Else
    ''    Call NGSAuditAdd(xEmpID, xUptType, xForm, lStr("Other Date 1"), xDate1, xDate2, xLDate)
    ''End If
    Call NGSAuditAdd(xEmpID, xUptType, xForm, lStr("Other Date 1"), xDate1, xDate2, xLDate)
End If
'NGS Start Date change - end
'NGS End Date change - begin
If IsDate(oOTHERDATE2) Then
    xDate1 = CVDate(oOTHERDATE2)
Else
    xDate1 = ""
End If
If IsDate(dlpDate(25).Text) Then
    xDate2 = CVDate(dlpDate(25).Text)
Else
    xDate2 = ""
End If
If Not (xDate1 = xDate2) Then
    If IsDate(xDate2) Then
        If CVDate(xDate2) > CVDate(Date) Then
            xLDate = CVDate(xDate2)
        End If
    End If

    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Other Date 2"), xDate1, xDate2, xLDate)
End If
'NGS End Date change - end

If Not (SavPT = clpPT) Then
    Call NGSAuditAdd(glbLEE_ID, xUptType, "Status/Dates", lStr("Category"), SavPT, clpPT.Text, xLDate)
End If
If Not (OVadim11 = clpVadim1.Text) Then '"NGS Sub Group"
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Vadim Field 1"), OVadim11, clpVadim1.Text, xLDate)
End If '
If Not (OVadim21 = clpVadim2.Text) Then '"Pay Group"
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Vadim Field 2"), OVadim21, clpVadim2.Text, xLDate)
End If
'Hire Date change - begin
If IsDate(SavDOH) Then
    xDate1 = CVDate(SavDOH)
Else
    xDate1 = ""
End If
If IsDate(dlpDate(7).Text) Then
    xDate2 = CVDate(dlpDate(7).Text)
Else
    xDate2 = ""
End If
If Not (xDate1 = xDate2) Then
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", "Original Hire Date", xDate1, xDate2, xLDate)
End If
'Hire Date change - end
'Date Of Death change - begin
If IsDate(oPENSIONDATE5) Then
    xDate1 = CVDate(oPENSIONDATE5)
Else
    xDate1 = ""
End If
If IsDate(dlpDate(22).Text) Then
    xDate2 = CVDate(dlpDate(22).Text)
Else
    xDate2 = ""
End If
If Not (xDate1 = xDate2) Then
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Pension Date 5"), xDate1, xDate2, xLDate)
End If
'Date Of Death change - end
'Retire Date change - begin
If IsDate(oPENSIONDATE6) Then
    xDate1 = CVDate(oPENSIONDATE6)
Else
    xDate1 = ""
End If
If IsDate(dlpDate(23).Text) Then
    xDate2 = CVDate(dlpDate(23).Text)
Else
    xDate2 = ""
End If
If Not (xDate1 = xDate2) Then
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", lStr("Pension Date 6"), xDate1, xDate2, xLDate)
End If
'Retire Date change - end

'Ticket #22718 Franks 12/12/2012
'Spouse works at Woodbridge - Begin
If OWITHSPOUSE <> chkSpouse Then
    If OWITHSPOUSE Then xOldNGSVal = "Y" Else xOldNGSVal = "N"
    If chkSpouse Then xNewNGSVal = "Y" Else xNewNGSVal = "N"
    Call NGSAuditAdd(xEmpID, xUptType, "Status/Dates", "Spouse works at WFC", xOldNGSVal, xNewNGSVal, xLDate)
End If
'Spouse works at Woodbridge - end

AUDIT_NGS_TRANS = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING NGS AUDIT RECORD", "NGS AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub AUDIT_GWL_TRANS()
Dim xEmpID
Dim xForm As String
Dim xTranType
Dim xChgType
Dim xEDate
Dim xLDate
Dim uptFlag As Boolean
On Error GoTo AUDIT_ERR

    If Not glbIsGWL Then Exit Sub
    '"   Only transfer employees who have a Benefit Group Code on their Status/Dates screen.
    If Len(clpBGroup.Text) = 0 Then Exit Sub
    'If NewHireForms.count = 0 Then Exit Sub 'new hire
    
    If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
        xEmpID = glbTERM_ID
    Else
        xEmpID = glbLEE_ID
    End If
    
    uptFlag = False
    If NewHireForms.count > 0 Then
        xTranType = "A"
        xChgType = "New Hire"
        uptFlag = True
    Else
        'check the change
        If Not (OBenGrp = clpBGroup.Text) Then
            xTranType = "R"
            'xChgType = "Personal Info"
            xChgType = "Group Change"
        End If
    End If
    
    xEDate = dlpDate(7).Text
    xForm = "Status/Dates"
    xLDate = Date
    If IsDate(xEDate) Then
        If CVDate(xEDate) > CVDate(xLDate) Then
            xLDate = xEDate
        End If
    End If
    If uptFlag Then
        Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Hire Date", "", dlpDate(7).Text, xLDate)
    End If
    
    Exit Sub

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING GWL AUDIT RECORD", "GWL AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Private Function AUDIT_MANULIFE_TRANS()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsBene As New ADODB.Recordset
Dim rsDepend As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim SQLQ As String
Dim xBenTermDate
Dim xEmpStatCode
Dim xUptFlag As Boolean

On Error GoTo AUDIT_ERR
AUDIT_MANULIFE_TRANS = False
    
'Manulife Audit
'with Certificate#
If Len(txtUserText1.Text) = 0 Then Exit Function
If Len(txtUserText2.Text) = 0 Then Exit Function
If Len(txtUserNum1.Text) = 0 Then Exit Function

xUptFlag = False
xEmpStatCode = clpCode(1).Text
If xEmpStatCode = "FS" Then
    'Emp Status was changed to "FS"
    'The Benefits will be terminated after 1 year since From Date
    If Not IsDate(dlpDate(15).Text) Then Exit Function
    If SavEmp = xEmpStatCode And oFDate = dlpDate(15).Text Then Exit Function    'No Status change
    xBenTermDate = DateAdd("yyyy", 1, dlpDate(15).Text)
    xUptFlag = True
End If
If xEmpStatCode = "SALC" Then
    'The Benefits will be terminated on To Date
    If Not IsDate(dlpDate(16).Text) Then Exit Function
    If SavEmp = xEmpStatCode And OTDate = dlpDate(16).Text Then Exit Function
    xBenTermDate = dlpDate(16).Text
    xUptFlag = True
End If
If xEmpStatCode = "RET" Then
    'The Benefits will be terminated on From Date
    If Not IsDate(dlpDate(15).Text) Then Exit Function
    If SavEmp = xEmpStatCode And oFDate = dlpDate(15).Text Then Exit Function
    xBenTermDate = dlpDate(15).Text
    xUptFlag = True
End If

If Not xUptFlag Then Exit Function


rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    rsTB.Close:    GoTo MODNOUPD_Ben
End If


'Benefits
SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & "AND (BF_BCODE = 'DENT' OR BF_BCODE = 'EHC' OR BF_BCODE = 'HCSA' OR BF_BCODE = 'HCSA1') "
rsBene.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsBene.EOF Then
    rsBene.Close
    GoTo MODNOUPD_Ben
End If


Do While Not rsBene.EOF
    If Len(rsBene("BF_POLICY")) > 0 Then
        If Not IsDate(rsBene("BF_CEASEDATE")) Then 'No Benefit End Date
            If rsTA.State <> 0 Then rsTA.Close
            rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            
            rsTA.AddNew
            rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
            rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
            rsTA("MT_PT_TABL") = "EDPT"
            rsTA("MT_TYPE") = "T"
            rsTA("MT_BENEFIT") = rsBene("BF_BCODE")
            rsTA("MT_EDATE") = rsBene("BF_EDATE")
            rsTA("MT_CEASEDATE") = xBenTermDate
            rsTA("MT_COVER") = rsBene("BF_COVER")
            rsTA("MT_COMPNO") = "001"
            rsTA("MT_EMPNBR") = glbLEE_ID
            rsTA("MT_POLICY_NO") = rsBene("BF_POLICY")
            If Len(txtUserNum1.Text) > 0 And IsNumeric((txtUserNum1.Text)) Then rsTA("MT_ACCOUNT_NO") = txtUserNum1.Text
            If Len(txtUserText1.Text) > 0 Then rsTA("MT_CERT_NO") = txtUserText1.Text
            If Len(txtUserText2.Text) > 0 Then rsTA("MT_COVERAGE_CLASS") = txtUserText2.Text
            rsTA("MT_EMP") = clpCode(1).Text
            If Len(SavEmp) > 0 Then rsTA("MT_OLDEMP") = SavEmp
            rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
            rsTA("MT_UPLOAD") = "N" '
            rsTA("MT_LUSER") = glbUserID
            rsTA("MT_LDATE") = Format(xBenTermDate, "SHORT DATE")
            rsTA("MT_LTIME") = Time$
            rsTA.Update
            
            rsBene("BF_CEASEDATE") = xBenTermDate
            rsBene.Update
        End If
    End If
    rsBene.MoveNext
Loop
rsBene.Close

MODNOUPD_Ben:


AUDIT_MANULIFE_TRANS = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT RECORD", "MANULIFE AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function HOOPPBenFlag(xEmpNo)
Dim rsHOOPP As New ADODB.Recordset
Dim SQLQ As String
    HOOPPBenFlag = False
    SQLQ = "SELECT ED_EMPNBR, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsHOOPP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHOOPP.EOF Then
        If Not IsNull(rsHOOPP("ED_REGION")) Then
            If rsHOOPP("ED_REGION") = "2" Or rsHOOPP("ED_REGION") = "3" Or rsHOOPP("ED_REGION") = "4" Then
                HOOPPBenFlag = True
            End If
        End If
    End If
    rsHOOPP.Close
End Function

Private Function Popup_JobCode()
    Dim xJobCode As String
    Dim rsHRJOB As New ADODB.Recordset
    Dim SQLQ As String
    
    glbJob = ""
    frmNewEmployee.Show 1
    xJobCode = glbJob

    'Create a new record in the HR_JOB_HISTORY table for the Staging Career table to pick up
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE 1 = 2"
    rsHRJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsHRJOB.AddNew
    rsHRJOB("JH_COMPNO") = "001"
    rsHRJOB("JH_EMPNBR") = glbLEE_ID
    rsHRJOB("JH_SDATE") = dlpDate(7).Text
    rsHRJOB("JH_CURRENT") = "1"
    rsHRJOB("JH_JOB") = xJobCode
    rsHRJOB("JH_JREASON") = "NEWH"
    rsHRJOB("JH_GLNO") = rsDATA("ED_GLNO")

    rsHRJOB("JH_SHIFT") = clpCode(2).Text  'Ticket #14791 - as default value for new record
    rsHRJOB("JH_DIV") = rsDATA("ED_DIV")
    rsHRJOB("JH_DEPTNO") = rsDATA("ED_DEPTNO")
    rsHRJOB("JH_EMP") = clpCode(1).Text
    rsHRJOB("JH_ORG") = clpCode(2).Text
    rsHRJOB("JH_PT") = clpPT.Text
    rsHRJOB("JH_SECTION") = rsDATA("ED_SECTION")
    
    rsHRJOB("JH_LDATE") = Date
    rsHRJOB("JH_LTIME") = Time$
    rsHRJOB("JH_LUSER") = glbUserID
    rsHRJOB.Update
    rsHRJOB.Close
    Set rsHRJOB = Nothing
End Function

Private Sub TabStringSetup()
    tabDates.Height = 2600
    fraDateEmp.Height = 2050
    fraDateEmp.Width = 9000
    fraDatePension.Height = 2050
    fraDatePension.Width = 9000
    fraDateOther.Height = 2050
    fraDateOther.Width = 9000

    fraDateEmp.Top = tabDates.Top + 500
    fraDatePension.Top = tabDates.Top + 500
    fraDateOther.Top = tabDates.Top + 500
    fraDateEmp.Left = tabDates.Left + 100
    fraDatePension.Left = tabDates.Left + 100
    fraDateOther.Left = tabDates.Left + 100
    fraDateEmp.BorderStyle = 0: fraDatePension.BorderStyle = 0: fraDateOther.BorderStyle = 0
    tabDates.Visible = True
    
    'Ticket #17919 - So that on employee scroll it does not display the frame if the user do not have access to
    'birth date
    'fraDateEmp.Visible = True
    fraDatePension.Visible = gSec_Show_DOB
    fraDateEmp.Visible = False
    fraDateOther.Visible = False
    Call tabDates_Click
    
End Sub

Private Sub LabelsSetup()

Call setCaption(lblUnion)
lblUnionEDate.Caption = lblUnion.Caption & " Effective"     'Ticket #29230 - New Field

'This issue is only for County of Essex
'on second and subsequent load of the screen it puts the Last Hire field label as "Recent Hire"
'instead of "Original Hire" as Original Hire Date field label has been relabeled to "Recent Hire"
'so when the program finds Last Hire as "Original Hire" it reads in the table the relabeled
'value for Original Hire Date field.
lblLHire.Caption = "Last Hire"
lblFDay.Caption = "First Day"
lblLDay.Caption = "Last Day"
lblPT.Caption = "Category"
'Friesens has seniority and doh
lblOHire.Caption = "Original Hire"
lblSen.Caption = "Seniority"
lblElig.Caption = "Eligibility"

lblEEType.Caption = "Employment Type"
lblIPhone.Caption = "Internal Phone Extension"
lblEmail.Caption = "Email Address"
lblBen.Caption = "Benefit Group"

lblOHire.Caption = "Original Hire"
lblSen.Caption = "Seniority"
lblPT.Caption = "Category"
lblLHire.Caption = "Last Hire"
lblUDate.Caption = "Union Date"
lblFDay.Caption = "First Day"
lblLDay.Caption = "Last Day"
lblODate.Caption = "OMERS Date"
lblUDay.Caption = "User Defined"
lblDeptStart.Caption = "Depart. Start Date"
lblDivStart.Caption = "Division Start Date"
lblElig.Caption = "Eligibility"
lblEarlR.Caption = "Earliest Retirement"
lblNorR.Caption = "Normal Retirement"
lblLateR.Caption = "Latest Retirement"
'lblTitle(20).Caption = ""
lblHireCode.Caption = "Hire Code"
lblSalDist.Caption = "Salary Distribution"
lblUserText1.Caption = "User Text 1"
lblUserText2.Caption = "User Text 2"
lblUserNum1.Caption = "User Number 1"
lblUserNum2.Caption = "User Number 2"
Call setCaption(lblOHire)
Call setCaption(lblSen)
Call setCaption(lblPT)
lblPTEDate.Caption = lblPT.Caption & " Effective"
Call setCaption(lblLHire)
Call setCaption(lblUDate)
Call setCaption(lblFDay)
Call setCaption(lblLDay)
Call setCaption(lblODate)
Call setCaption(lblUDay)
Call setCaption(lblDeptStart)
Call setCaption(lblDivStart)
Call setCaption(lblElig)
Call setCaption(lblEarlR)
Call setCaption(lblNorR)
Call setCaption(lblLateR)
Call setCaption(lblHireCode)
Call setCaption(lblSalDist)
Call setCaption(lblUserText1)
Call setCaption(lblUserText2)
Call setCaption(lblUserNum1)
Call setCaption(lblUserNum2)
txtUserText1.Tag = lStr(txtUserText1.Tag)
txtUserText2.Tag = lStr(txtUserText2.Tag)
txtUserNum1.Tag = lStr(txtUserNum1.Tag)
txtUserNum2.Tag = lStr(txtUserNum2.Tag)

'New label fields for v7.8 Ticket #15576 - Begin
lblPenDate(0).Caption = "Pension Date 1"
lblPenDate(1).Caption = "Pension Date 2"
lblPenDate(2).Caption = "Pension Date 3"
lblPenDate(3).Caption = "Pension Date 4"
lblPenDate(4).Caption = "Pension Date 5"
lblPenDate(5).Caption = "Pension Date 6"
lbOtherDate(0).Caption = "Other Date 1"
lbOtherDate(1).Caption = "Other Date 2"
lbOtherDate(2).Caption = "Other Date 3"
lbOtherDate(3).Caption = "Other Date 4"
lbOtherDate(4).Caption = "Other Date 5"
lbOtherDate(5).Caption = "Other Date 6"
lbOtherDate(6).Caption = "Other Date 7"
lbOtherDate(7).Caption = "Other Date 8"
lbOtherDate(9).Caption = "Other Date 9"
lbOtherDate(9).Caption = "Other Date 10"

Call setCaption(lblEEType)
Call setCaption(lblIPhone)
Call setCaption(lblEmail)
Call setCaption(lblBen)
Call setCaption(lblPenDate(0))
Call setCaption(lblPenDate(1))
Call setCaption(lblPenDate(2))
Call setCaption(lblPenDate(3))
Call setCaption(lblPenDate(4))
Call setCaption(lblPenDate(5))
Call setCaption(lbOtherDate(0))
Call setCaption(lbOtherDate(1))
Call setCaption(lbOtherDate(2))
Call setCaption(lbOtherDate(3))
Call setCaption(lbOtherDate(4))
Call setCaption(lbOtherDate(5))
Call setCaption(lbOtherDate(6))
Call setCaption(lbOtherDate(7))
Call setCaption(lbOtherDate(8))
Call setCaption(lbOtherDate(9))
'New label fields for v7.8 Ticket #15576 - End

lblTitle(20).Caption = lStr("Location")

End Sub

Private Function AUDITSTAT2()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xDiv, xPT
Dim x
Dim UpdateAudit As Boolean
On Error GoTo AUDIT_ERR
AUDITSTAT2 = False

xADD = False

If oHireCode <> clpCode(6).Text Then UpdateAudit = True
If oPENSIONDATE1 <> dlpDate(18).Text Then UpdateAudit = True
If oPENSIONDATE2 <> dlpDate(19).Text Then UpdateAudit = True
If oPENSIONDATE3 <> dlpDate(20).Text Then UpdateAudit = True
If oPENSIONDATE4 <> dlpDate(21).Text Then UpdateAudit = True
If oPENSIONDATE5 <> dlpDate(22).Text Then UpdateAudit = True
If oPENSIONDATE6 <> dlpDate(23).Text Then UpdateAudit = True
If oOTHERDATE1 <> dlpDate(24).Text Then UpdateAudit = True
If oOTHERDATE2 <> dlpDate(25).Text Then UpdateAudit = True
If oOTHERDATE3 <> dlpDate(26).Text Then UpdateAudit = True
If oOTHERDATE4 <> dlpDate(27).Text Then UpdateAudit = True
If oOTHERDATE5 <> dlpDate(28).Text Then UpdateAudit = True
If oOTHERDATE6 <> dlpDate(29).Text Then UpdateAudit = True
If oOTHERDATE7 <> dlpDate(30).Text Then UpdateAudit = True
If oOTHERDATE8 <> dlpDate(31).Text Then UpdateAudit = True
If oOTHERDATE9 <> dlpDate(32).Text Then UpdateAudit = True
If oOTHERDATE10 <> dlpDate(33).Text Then UpdateAudit = True

If UpdateAudit Then
    GoTo MODUPD
Else
    GoTo MODNOUPD
End If


MODUPD:
    If rsTA.State <> 0 Then rsTA.Close
    rsTA.Open "SELECT * FROM HRAUDIT2 WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

    rsTB.Open "select ED_DIV,ED_PT,ED_PAYROLL_ID,ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    
    If Not rsTB.EOF Then
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_HIRECODE_TABL") = "EDHC": rsTA("AU_ORG_TABL") = "EDOR"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_DIVUPL") = xDiv
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    If Not IsNull(rsTB("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTB("ED_PAYROLL_ID")
    If Not IsNull(rsTB("ED_SECTION")) Then rsTA("AU_SECTION") = rsTB("ED_SECTION")
    If oHireCode <> clpCode(6).Text Then rsTA("AU_HIRECODE") = clpCode(6).Text
    If oPENSIONDATE1 <> dlpDate(18).Text Then
        If IsDate(dlpDate(18).Text) Then rsTA("AU_PENSIONDATE1") = dlpDate(18).Text
    End If
    If oPENSIONDATE2 <> dlpDate(19).Text Then
        If IsDate(dlpDate(19).Text) Then rsTA("AU_PENSIONDATE2") = dlpDate(19).Text
    End If
    If oPENSIONDATE3 <> dlpDate(20).Text Then
        If IsDate(dlpDate(20).Text) Then rsTA("AU_PENSIONDATE3") = dlpDate(20).Text
    End If
    If oPENSIONDATE4 <> dlpDate(21).Text Then
        If IsDate(dlpDate(21).Text) Then rsTA("AU_PENSIONDATE4") = dlpDate(21).Text
    End If
    If oPENSIONDATE5 <> dlpDate(22).Text Then
        If IsDate(dlpDate(22).Text) Then rsTA("AU_PENSIONDATE5") = dlpDate(22).Text
    End If
    If oPENSIONDATE6 <> dlpDate(23).Text Then
        If IsDate(dlpDate(23).Text) Then rsTA("AU_PENSIONDATE6") = dlpDate(23).Text
    End If
    
    If oOTHERDATE1 <> dlpDate(24).Text Then
        If IsDate(dlpDate(24).Text) Then rsTA("AU_OTHERDATE1") = dlpDate(24).Text
    End If
    If oOTHERDATE2 <> dlpDate(25).Text Then
        If IsDate(dlpDate(25).Text) Then rsTA("AU_OTHERDATE2") = dlpDate(25).Text
    End If
    If oOTHERDATE3 <> dlpDate(26).Text Then
        If IsDate(dlpDate(26).Text) Then rsTA("AU_OTHERDATE3") = dlpDate(26).Text
    End If
    If oOTHERDATE4 <> dlpDate(27).Text Then
        If IsDate(dlpDate(27).Text) Then rsTA("AU_OTHERDATE4") = dlpDate(27).Text
    End If
    If oOTHERDATE5 <> dlpDate(28).Text Then
        If IsDate(dlpDate(28).Text) Then rsTA("AU_OTHERDATE5") = dlpDate(28).Text
    End If
    If oOTHERDATE6 <> dlpDate(29).Text Then
        If IsDate(dlpDate(29).Text) Then rsTA("AU_OTHERDATE6") = dlpDate(29).Text
    End If
    If oOTHERDATE7 <> dlpDate(30).Text Then
        If IsDate(dlpDate(30).Text) Then rsTA("AU_OTHERDATE7") = dlpDate(30).Text
    End If
    If oOTHERDATE8 <> dlpDate(31).Text Then
        If IsDate(dlpDate(31).Text) Then rsTA("AU_OTHERDATE8") = dlpDate(31).Text
    End If
    If oOTHERDATE9 <> dlpDate(32).Text Then
        If IsDate(dlpDate(32).Text) Then rsTA("AU_OTHERDATE9") = dlpDate(32).Text
    End If
    If oOTHERDATE10 <> dlpDate(33).Text Then
        If IsDate(dlpDate(33).Text) Then rsTA("AU_OTHERDATE10") = dlpDate(33).Text
    End If
    rsTA("AU_LDATE") = Date
    If IsDate(dlpDate(7).Text) Then 'Ticket #21786 Franks 03/23/2012
        If CVDate(dlpDate(7).Text) > Date Then
            rsTA("AU_LDATE") = dlpDate(7).Text
        End If
    End If
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA.Update
    rsTA.Close
    rsTB.Close
    
MODNOUPD:
AUDITSTAT2 = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT2 RECORD", "AUDIT2 FILE", "UPDATE")
Call RollBack
Resume Next
End Function

Private Sub txtUserText2_Change()
If glbWFC Then 'Ticket #22448 Franks
'    comUserText2.Text = txtUserText2.Text
    comUserText2.ListIndex = FindCBIndex(comUserText2, Left((txtUserText2 & " - "), 4), 4)
'Ticket #24976 - VitalAire Canada Inc.
ElseIf glbCompSerial = "S/N - 2380W" Then
    comUserText2.ListIndex = -1
    comUserText2.ListIndex = GetUserText2Index(Trim(txtUserText2.Text))
End If
End Sub

Private Sub PopcomUserText2(xBenAccount) 'Ticket #22448 Franks
Dim rsBenSetup As New ADODB.Recordset
Dim SQLQ As String
Dim xCurVal
    'xCurVal = comUserText2.Text
    comUserText2.Clear
    If Len(xBenAccount) > 0 Then
        SQLQ = "SELECT * FROM WFC_BENEFIT_ACCOUNT_SETUP WHERE BU_BEN_ACCOUNT = '" & xBenAccount & "' ORDER BY BU_CLASS"
        If rsBenSetup.State <> 0 Then rsBenSetup.Close
        rsBenSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsBenSetup.EOF
            comUserText2.AddItem rsBenSetup("BU_CLASS") & " - " & rsBenSetup("BU_CLASS_DESC")
            rsBenSetup.MoveNext
        Loop
        rsBenSetup.Close
    End If
    Call txtUserText2_Change
    'comUserText2.Text = xCurVal
End Sub

Private Sub txtUserText2_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserText2_LostFocus()
    If glbWFC Then
        If Trim(Len(txtUserText2)) > 0 Then
            txtUserText2.Text = UCase(txtUserText2)
        End If
    End If
End Sub

Private Sub SamuelFieldsLayout()
    'Vadim Field 1
    lblVadim11.Width = 1875 'Ticket #23448  lblUserNum2.Width
    lblVadim11.Left = 5670 'Ticket #23448 lblUserNum2.Left
    lblVadim11.Top = txtUserNum1.Top + 285 + 20 + 60
    clpVadim1.Left = clpBGroup.Left
    clpVadim1.Top = txtUserNum1.Top + 285 + 60
    lblVadim11.Visible = True
    clpVadim1.Visible = True
    clpVadim1.DataField = "ED_VADIM1"
    lblVadim11.Caption = lStr("Vadim Field 1")
    
    'Supervisor Code
    lblSupervisor.Left = lblUserNum1.Left
    lblSupervisor.Top = lblVadim11.Top
    clpCode(8).Left = clpCode(6).Left
    clpCode(8).Top = clpVadim1.Top
    lblSupervisor.Visible = True
    clpCode(8).Visible = True
    clpCode(8).DataField = "ED_SUPCODE"
    lblSupervisor.Caption = lStr("Supervisor Code")
    
    'Release 8.0 - The Email Import button is overlapping this so I am moving to Hire Code level
    'lblImport.Top = lblEmail.Top
    'imgNoSec.Top = lblEmail.Top
    'imgSec.Top = lblEmail.Top
    'cmdImport.Top = txtEmail.Top

    lblImport.Top = lblHireCode.Top
    imgNoSec.Top = lblHireCode.Top
    imgSec.Top = lblHireCode.Top
    cmdImport.Top = lblHireCode.Top - 80

    
End Sub

Private Sub WFCVadimFieldsLayout()
    lblUserNum2.Visible = False
    txtUserNum2.Visible = False
    lblUserText2.Left = lblUserText1.Left
    lblUserText2.Top = 3105 'lblSpouse.Top
    lblUserText2.Alignment = 0
    cmdEditUserText2.Left = cmdEditUserText1.Left
    cmdEditUserText2.Top = 3105 'lblSpouse.Top
    txtUserText2.Left = txtUserText1.Left
    txtUserText2.Top = 3105 'lblSpouse.Top
    'Ticket #22448 Franks - begin
    comUserText2.Left = txtUserText1.Left ' + 2000
    comUserText2.Width = 3000
    comUserText2.Top = 3105 - 50 ' lblSpouse.Top - 50
    comUserText2.Visible = True
    txtUserText2.Visible = False
    'Ticket #22448 Franks - end
    lblVadim11.Left = 5300 'lblBen.Left
    lblVadim11.Top = txtUserNum1.Top + 20
    lblVadim21.Left = 5300 'lblBen.Left
    lblVadim21.Top = txtUserText1.Top + 20 '
    clpVadim1.Enabled = False 'Ticket #21544 Franks 02/07/2012
    clpVadim1.Left = clpBGroup.Left
    clpVadim1.Top = txtUserNum1.Top
    clpVadim2.Left = clpBGroup.Left
    clpVadim2.Top = txtUserText1.Top '
    cmdEditNGSSub.Left = clpVadim1.Left - 650
    cmdEditNGSSub.Top = clpVadim1.Top
    cmdEditNGSSub.Visible = True
    lblVadim11.Visible = True
    lblVadim21.Visible = True
    clpVadim1.Visible = True
    clpVadim2.Visible = True
    clpVadim1.DataField = "ED_VADIM1"
    clpVadim2.DataField = "ED_VADIM2"
    'Label master
    lblVadim11.Caption = lStr("Vadim Field 1")
    lblVadim21.Caption = lStr("Vadim Field 2")
    'tabindes
    txtUserNum1.TabIndex = txtUserText1.TabIndex + 1
    txtUserText2.TabIndex = txtUserText1.TabIndex + 2
    clpVadim2.TabIndex = txtUserText1.TabIndex + 3
    clpVadim1.TabIndex = txtUserText1.TabIndex + 4

    'Ticket #22718 Franks 12/12/12
    'Turn on the field "Spouse works at Linamar". Change Linamar to Woodbridge
    lblSpouse.Top = clpCode(6).Top 'Hire Code
    chkSpouse.Top = clpCode(6).Top 'Hire Code
    lblSpouse.Left = 5300
    chkSpouse.Left = 8400
    lblSpouse.Alignment = 1
    lblSpouse.Width = 3000
    lblSpouse.Caption = "Spouse works at Woodbridge"
    lblSpouse.Visible = True
    chkSpouse.Visible = True

End Sub
Private Sub DispPayGroup(xCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
    If glbEmpCountry = "U.S.A." Then 'Ticket #20387 Franks 05/27/2011
        If Len(xCode) > 0 Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_SUB_GROUP = '" & xCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                clpVadim2.Text = rsTemp("NG_PAY_GROUP")
            End If
            rsTemp.Close
        End If
    End If
End Sub
Private Sub DispNGSSubGroup(xCode) 'Ticket #19266
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
    'If Len(clpVadim1.Text) = 0 And Len(xCode) > 0 Then
    If glbEmpCountry = "U.S.A." Then 'Ticket #20387 Franks 05/27/2011
        If Len(xCode) > 0 Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_PAY_GROUP = '" & xCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                clpVadim1.Text = rsTemp("NG_SUB_GROUP")
            End If
            rsTemp.Close
        End If
    End If
End Sub

Private Sub DispNGSGroups() 'Ticket #19306
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xLocOrg
    
    If Not glbNGS_OnFlag Then Exit Sub
    
    If Not glbEmpCountry = "U.S.A." Then  'Ticket #20387 Franks 05/27/2011
        Exit Sub
    End If
    
    'Ticket #25352 Franks 04/16/2014 - exclude "COOP" and "STUD"
    If (clpCode(1).Text = "COOP" Or clpCode(1).Text = "STUD") Then
        Exit Sub
    End If
    
    'If Not clpPT.Text = "FT" Then 'Ticket #20638 Franks 07/18/2011
    If Not (clpPT.Text = "FT" Or clpPT.Text = "PT") Then 'Ticket #22991 Franks 12/24/2012
        If NewHireForms.count = 0 Then 'Ticket #23916 Franks for change only
            clpVadim1.Text = "" 'Ticket #20712 Franks 08/03/2011
        End If
        Exit Sub
    End If
    
    'If Not (SavOrg = clpCode(2).Text) Then
    'Franks 10/16/2012 -  add status and clpPT here
    If Not (SavOrg = clpCode(2).Text) Or Not (SavEmp = clpCode(1).Text) Or Not (SavPT = clpPT) Then
        If Len(clpCode(2).Text) = 0 Then 'remove Union
            If NewHireForms.count = 0 Then 'Ticket #23916 Franks for change only
                clpVadim1.Text = ""
                'clpVadim2.Text = ""
            End If
        Else
            'Ticket #20305 Franks 05/17/2011 -  add status as a key
            'with Status code
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsDATA("ED_DIV") & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & clpCode(2).Text & "' "
            SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & clpCode(1).Text & "' "
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            
            If rsTemp.EOF Then 'Ticket #23564 Franks 04/17/2013
            'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" then compare ED_EMP with not equal to
                SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsDATA("ED_DIV") & "' "
                SQLQ = SQLQ & "AND NG_ORG = '" & clpCode(2).Text & "' "
                SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
                SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & clpCode(1).Text & "') " 'convert "-ACT2" to "ACT2"; no "-" then ""
                If rsTemp.State <> 0 Then rsTemp.Close
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                'if not found then without Status code
                If rsTemp.EOF Then
                    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsDATA("ED_DIV") & "' "
                    SQLQ = SQLQ & "AND NG_ORG = '" & clpCode(2).Text & "' "
                    SQLQ = SQLQ & "AND ((NG_PLAN_CODE IS NULL) OR NOT( NG_PLAN_CODE ='" & clpCode(1).Text & "')) "
                    If rsTemp.State <> 0 Then rsTemp.Close
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                End If
            End If
            
            If Not rsTemp.EOF Then
                clpVadim2.Text = rsTemp("NG_PAY_GROUP")
                clpVadim1.Text = rsTemp("NG_SUB_GROUP")
                If Not IsNull(rsTemp("NG_BENEFIT_GROUP")) Then 'Ticket #23903 Franks 06/20/2013
                    clpBGroup.Text = rsTemp("NG_BENEFIT_GROUP")
                Else
                    clpBGroup.Text = ""
                End If
            Else
                clpVadim1.Text = ""
                'clpVadim2.Text = ""
                clpBGroup.Text = "" 'Ticket #23903 Franks 06/20/2013
            End If
            rsTemp.Close
        End If
        'If Len(clpCode(2).Text) >= 0 And Len(clpVadim1.Text) > 0 Then
        '    If Len(SavOrg) > 0 Then 'for change only, not for new hire
        '        'Union Code Change
        '        'Display a pop-up asking for "Other Date 1" - NGS Effective Date
        '        tabDates.SelectedItem = tabDates.Tabs(3)
        '        glbChgTermDate = "" 'dlpDate(24).Text
        '        frmMsgTerm.PenTermDate = "NGS_EffectiveDate"
        '        frmMsgTerm.Show 1
        '        'tabDates.SelectedItem = tabDates.Tabs(3)
        '        dlpDate(24).Text = glbChgTermDate
        '        dlpDate(24).SetFocus
        '    End If
        'End If
    End If
End Sub

'Private Sub NGSDateWindow(xType, xUnion)
Private Sub WFC_NGS_Trans()
Dim xUnion As String
Dim xType As String
Dim xSalHly As String
Dim xMsg
    
    glbMsgCustomVal = 0
    
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    xUnion = clpCode(2).Text
    If xUnion = "NONE" Or xUnion = "EXEC" Then
        xSalHly = "Y"
    Else
        xSalHly = "N"
    End If
    xNGSpopFlag = True 'pop windowonly only show up once to ask Other Date 1
        
        
    'Ticket #23247 Franks 03/12/2014 - begin
    'from FT to non PT(such as other), the program removed ngs sub group, to we have to do the transaction here
    If SavPT = "FT" Then
        If Not (clpPT.Text = "FT" Or clpPT.Text = "PT") Then  'Other
            If IsDate(dlpDate(24).Text) Then 'NGS Start Date
                glbMsgCustomVal = 4
                If Len(dlpDate(25).Text) = 0 Then 'No NGS End Date
                    glbChgTermDate = "" ' dlpDate(25).Text
                    frmMsgTerm.PenTermDate = "NGS_TermDate"
                    frmMsgTerm.Show 1
                    If Len(glbChgTermDate) > 0 Then
                        dlpDate(25).Text = glbChgTermDate
                    End If
                    tabDates.Tabs(3).selected = True
                    Call tabDates_Click
                    dlpDate(25).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    'Ticket #23247 Franks 03/12/2014 - end
    
    'No NGS Sub Gropu, skip
    If Len(clpVadim1.Text) = 0 Then
        Exit Sub
    End If
    
    ''xUnion = clpCode(2).Text
    ''If xUnion = "NONE" Or xUnion = "EXEC" Then
    ''    xSalHly = "Y"
    ''Else
    ''    xSalHly = "N"
    ''End If
    ''xNGSpopFlag = True 'pop windowonly only show up once to ask Other Date 1
    
    If NewHireForms.count > 0 Then
    'New hire: For Salaried employees, the "Other Date 1" defaults to the Original Date of Hire.
        If xSalHly = "Y" Then
            If IsDate(dlpDate(7).Text) Then
                If Len(dlpDate(24).Text) = 0 Then
                    If glbTrsHourWeek > 0 And glbTrsHourWeek < 20 Then
                        'Ticket #25248 Franks 03/24/2014
                        'On New Hire (US NGS employees), if the hours per week is less than 20, don't update the NGS Start Date. Jerry
                    Else
                        dlpDate(24).Text = dlpDate(7).Text
                    End If
                End If
            End If
        End If
        Exit Sub
    End If
    
    
    'If xType = "FT/PT" Then
    If Not (SavPT = clpPT) Then
        'Change from: PT/SE/TR/OT to FT -
        'Display a pop-up asking for "Other Date 1".  For Salaried employees, the "Other Date 1" defaults to the Original Date of Hire. Salaried is defined if union code is equal to "NONE" or "EXEC". User will enter "Other Date 1" for hourly employees.
        If clpPT = "FT" Then
            tabDates.SelectedItem = tabDates.Tabs(3)
            If xSalHly = "Y" Then
                glbChgTermDate = dlpDate(7).Text
            Else
                glbChgTermDate = "" 'dlpDate(24).Text
                xNGSpopFlag = True
            End If
            If xNGSpopFlag Then 'pop windowonly only show up once
                xNGSpopFlag = False
                frmMsgTerm.PenTermDate = "NGS_EffectiveDate"
                frmMsgTerm.dlpTermDate = dlpDate(24).Text 'Ticket #23247 Franks 03/05/2014
                frmMsgTerm.Show 1
            End If
            'tabDates.SelectedItem = tabDates.Tabs(3)
            If Len(glbChgTermDate) > 0 Then
                dlpDate(24).Text = glbChgTermDate
                'If the Other Date 2(NGS End) is less than the Other Date 1(NGS Start), clear the date.
                If IsDate(dlpDate(25).Text) Then
                    If CVDate(dlpDate(25).Text) < CVDate(glbChgTermDate) Then
                        dlpDate(25).Text = ""
                    End If
                End If
            End If
            
            'Ticket #25178 Franks 03/11/2014 - begin
            glbMsgCustomVal = 7 'from PT to other(FT, ...)
            If clpPT.Text = "FT" Then 'employee from PT to FT. The NGS End Date should remove
                glbMsgCustomVal = 8
                If IsDate(dlpDate(25).Text) Then
                    dlpDate(25).Text = ""
                End If
            End If
            ''Ticket #25178 Franks 03/11/2014 - end
            dlpDate(24).SetFocus
        End If
        
        'Change from: FT to PT/SE/TR/OT
        'Display a pop-up asking for "Other Date 2".  The user must enter "Other Date 2".
        If Not (clpPT = "FT") Then
            'Ticket #23247 Franks 02/25/2014 - begin
            If clpPT.Text = "PT" Then
                xMsg = "Does this employee qualify for:"
                'xMsg = xMsg & " Will this LOA affect the Reporting Authority structures?"
                frmMsgYesNoUn.lblMsg.Caption = xMsg
                frmMsgYesNoUn.lblMsg.Alignment = 0
                Call frmMsgYesNoUn.WFCFrameSetup
                frmMsgYesNoUn.Show 1
                If glbMsgCustomVal = 4 Then 'If "No Benefits":
                    tabDates.SelectedItem = tabDates.Tabs(3)
                    glbChgTermDate = "" ' dlpDate(25).Text
                    frmMsgTerm.PenTermDate = "NGS_TermDate"
                    frmMsgTerm.Show 1
                    'tabDates.SelectedItem = tabDates.Tabs(3)
                    If Len(glbChgTermDate) > 0 Then
                        dlpDate(25).Text = glbChgTermDate
                    End If
                    dlpDate(25).SetFocus
                End If
                If glbMsgCustomVal = 5 Then 'If "Life Only":
                    'term company paid benefits except IE
                End If
                If glbMsgCustomVal = 6 Then
                    'If "All Benefits": no change
                End If
                'Ticket #23247 Franks 02/25/2014 - end
            Else 'non PT, FT and other...
                If Not clpPT.Text = "PT" Then 'from FT to non PT
                    tabDates.SelectedItem = tabDates.Tabs(3)
                    glbChgTermDate = "" ' dlpDate(25).Text
                    frmMsgTerm.PenTermDate = "NGS_TermDate"
                    frmMsgTerm.Show 1
                    'tabDates.SelectedItem = tabDates.Tabs(3)
                    If Len(glbChgTermDate) > 0 Then
                        dlpDate(25).Text = glbChgTermDate
                    End If
                    dlpDate(25).SetFocus
                    
                    'Ticket #23247 Franks 03/12/2014
                    'Change in Category - FT to a code not equal to PT
                    'use NGS End Date to be benefit end date
                    glbMsgCustomVal = 4
                End If
            End If
        End If
    End If
    
    'If xType = "Union" Then
    If Not (SavOrg = clpCode(2).Text) Then
        If xNGSpopFlag Then 'pop windowonly only show up once
            If Len(SavOrg) > 0 Then 'for change only, not for new hire
                If Len(dlpDate(24).Text) = 0 Then 'If NGS Start Date is blank
                    'Union Code Change
                    'Display a pop-up asking for "Other Date 1" - NGS Effective Date
                    tabDates.SelectedItem = tabDates.Tabs(3)
                    glbChgTermDate = "" 'dlpDate(24).Text
                    frmMsgTerm.PenTermDate = "NGS_EffectiveDate"
                    frmMsgTerm.Show 1
                    'tabDates.SelectedItem = tabDates.Tabs(3)
                    dlpDate(24).Text = glbChgTermDate
                    dlpDate(24).SetFocus
                End If
            End If
        End If
    End If
    
    'If xType = "DOH" Then
    If Len(SavDOH) > 0 Then 'Ticket #20360 Franks 05/20/2011
        If Not (CVDate(SavDOH) = CVDate(dlpDate(7).Text)) Then
            'If Salaried, update "Other Date 1" with new Original Hire Date.
            'For hourly employees, display a pop-up asking for "Other Date 1". (Optional).
            'If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
            If xSalHly = "Y" Then
                dlpDate(24).Text = dlpDate(7).Text
            Else
                If xNGSpopFlag Then 'pop windowonly only show up once
                    xNGSpopFlag = False
                    tabDates.SelectedItem = tabDates.Tabs(3)
                    glbChgTermDate = "" 'dlpDate(24).Text
                    frmMsgTerm.PenTermDate = "NGS_EffectiveDate"
                    frmMsgTerm.cmdCancelShow = True
                    frmMsgTerm.Show 1
                    'tabDates.SelectedItem = tabDates.Tabs(3)
                    If Len(glbChgTermDate) > 0 Then
                        dlpDate(24).Text = glbChgTermDate
                    End If
                    dlpDate(24).SetFocus
                End If
            End If
        End If
    End If
End Sub

Sub UpdateOxfordCurrentPosition()
Dim rsEmpJob As New ADODB.Recordset

rsEmpJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockOptimistic
If Not rsEmpJob.EOF Then
    rsEmpJob("JH_EMP") = clpCode(1).Text
    rsEmpJob("JH_ORG") = clpCode(2).Text
    rsEmpJob("JH_PT") = clpPT.Text
    rsEmpJob.Update
End If
rsEmpJob.Close
Set rsEmpJob = Nothing
End Sub

Private Sub UpdateCurrentPosition_KerrysPlace()
Dim rsEmpJob As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID
rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
Do While Not rsEmpJob.EOF
    'Matching the original value then only change
    If rsEmpJob("JH_ORG") = SavOrg Then
        rsEmpJob("JH_ORG") = clpCode(2).Text
    End If
    'rsEmpJob("JH_EMP") = clpCode(1).Text
    'rsEmpJob("JH_PT") = clpPT.Text
    rsEmpJob.Update
    
    rsEmpJob.MoveNext
Loop
rsEmpJob.Close
Set rsEmpJob = Nothing

End Sub

Private Function BenGroupExist(xBGroup)
Dim rsBGMST As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    If Len(xBGroup) > 0 Then
        SQLQ = "SELECT BM_BENEFIT_GROUP FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & xBGroup & "' "
        rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsBGMST.EOF Then
            retVal = True
        End If
        rsBGMST.Close
    End If
    BenGroupExist = retVal
End Function

Private Sub ComEUserTexts()

'Ticket #24976 - VitalAire Canada Inc.
If glbCompSerial = "S/N - 2380W" Then
    'User Text 1 - Payroll Types
    comUserText1.Clear
    comUserText1.AddItem "AL - AirLiquide"
    comUserText1.AddItem "FB - FT Salaried"
    comUserText1.AddItem "FFS - Fee For Service"
    
    'Ticket #25176 - changed
    'comUserText1.AddItem "HT - Hourly Temp"
    comUserText1.AddItem "HT - Hourly Timesheets"
    
    'Ticket #25176 - changed
    'comUserText1.AddItem "HTB - Hourly FT Benefits"
    comUserText1.AddItem "HTB - Hourly Timesheets Benefits"
    
    comUserText1.AddItem "HX - Hourly Fixed"
    comUserText1.AddItem "HXB - Hourly Fixed Benefits"
    
    'Ticket #25176 - New ones added
    comUserText1.AddItem "PPT - Permanent PT Timesheets"
    comUserText1.AddItem "PPX - Permanent PT Fixed"
    
    comUserText1.AddItem "UT - Union (GH)"
    comUserText1.AddItem "UV - Union (Campbell)"
    
    comUserText1.Left = txtUserText1.Left
    comUserText1.Top = txtUserText1.Top
    comUserText1.Width = 3350
    comUserText1.Visible = True
    txtUserText1.Visible = False
    
    'User Text 2 - Expansion Codes
    comUserText2.Clear
    comUserText2.AddItem "01 - Managers & Professionals"
    comUserText2.AddItem "02 - Admin, Driver, Fill Plant"
    comUserText2.AddItem "03 - Clinical"
    
    comUserText2.Left = txtUserText2.Left
    comUserText2.Top = txtUserText2.Top
    comUserText2.Width = 3000
    comUserText2.Visible = True
    txtUserText2.Visible = False
End If

End Sub

Private Sub ComEUserText1()

'Ticket #23537 - Essex Country Lib. - Remove this logic now
'Ticket #18789 Franks 05/06/2011
'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
'    comUserText1.Clear
'    comUserText1.AddItem ""
'    comUserText1.AddItem "7:30D0 - 7:30 DAILY OVERTIME (HQ)"
'    comUserText1.AddItem "8:30D0 - 8:30 DAILY OVERTIME (BR)"
'    comUserText1.AddItem "C34:30 - CUPE 34:30 WEEKLY OT(FT-BR)"
'    comUserText1.AddItem "C34:3Q - CUPE 34:30 WEEKLY OT(FTPT-HQ)"
'    comUserText1.AddItem "C37:00 - CUPE 37:00 WEEKLY OT(PT-BR)"
'    comUserText1.AddItem "OT-NON - Janitors & Desk Clk"
'    comUserText1.AddItem "N35:00 - Non-Union Salary "
'
'    comUserText1.Left = txtUserText1.Left
'    comUserText1.Top = txtUserText1.Top
'    comUserText1.Width = 3350
'    comUserText1.Visible = True
'End If

End Sub

Private Sub comUserText1_Click()

'Ticket #23537 - Essex Country Lib. - Remove this logic now
'Ticket #18789 Franks 05/06/2011
'If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
'    If comUserText1.ListIndex <> -1 Then
'        txtUserText1.Text = Trim(Left(comUserText1.Text, 6))
'    End If
'End If
'Ticket #24976 - VitalAire Canada Inc.
If glbCompSerial = "S/N - 2380W" Then
    If comUserText1.ListIndex <> -1 Then
        txtUserText1.Text = Trim(Left(comUserText1.Text, InStr(1, comUserText1.Text, "-") - 1))
    End If
End If
End Sub

Private Function GetUserText1Index(xEmpType) 'Ticket #18789 Franks 05/06/2011
'Ticket #23537 - Essex Country Lib. - Remove this logic now
Dim xIndex As Integer
'    xIndex = -1
'    If glbCompSerial = "S/N - 2296W" Then 'Essex County Library
'        Select Case Left(xEmpType, 6)
'        Case "": xIndex = 0
'        Case "7:30D0": xIndex = 1
'        Case "8:30D0": xIndex = 2
'        Case "C34:30": xIndex = 3
'        Case "C34:3Q": xIndex = 4
'        Case "C37:00": xIndex = 5
'        Case "OT-NON": xIndex = 6
'        Case "N35:00": xIndex = 7
'        End Select
'    End If
'    GetUserText1Index = xIndex
    'Ticket #24976 - VitalAire Canada Inc.
    xIndex = -1
    If glbCompSerial = "S/N - 2380W" Then
        Select Case xEmpType
            Case "AL": xIndex = 0
            Case "FB": xIndex = 1
            Case "FFS": xIndex = 2
            Case "HT": xIndex = 3
            Case "HTB": xIndex = 4
            Case "HX": xIndex = 5
            Case "HXB": xIndex = 6
            
            'Ticket #25176 - Added new ones
            Case "PPT": xIndex = 7
            Case "PPX": xIndex = 8
            
            Case "UT": xIndex = 9
            Case "UV": xIndex = 10
        End Select
    End If
    GetUserText1Index = xIndex
End Function

Private Function GetUserText2Index(xEmpType)
Dim xIndex As Integer
    xIndex = -1
    'Ticket #24976 - VitalAire Canada Inc.
    If glbCompSerial = "S/N - 2380W" Then
        Select Case xEmpType
            Case "01": xIndex = 0
            Case "02": xIndex = 1
            Case "03": xIndex = 2
        End Select
    End If
    GetUserText2Index = xIndex
End Function

Function CheckDupEmpEmail(xEmpnbr, xEmailAddr)
Dim RsSIN As New ADODB.Recordset
Dim xTerm_Emplist
Dim SQLQ
    CheckDupEmpEmail = False
    If Len(xEmailAddr) = 0 Then
        Exit Function
    End If
    SQLQ = "SELECT ED_EMPNBR,ED_EMAIL, ED_SIN,ED_SSN,ED_SURNAME,ED_FNAME FROM HREMP "
    SQLQ = SQLQ & "WHERE ED_EMAIL = '" & xEmailAddr & "' "
    SQLQ = SQLQ & "And ED_EMPNBR <> " & xEmpnbr
    RsSIN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsSIN.EOF Then
        CheckDupEmpEmail = True
    End If
    Dim xEmpList
    xEmpList = ""
    Do Until RsSIN.EOF
        xEmpList = xEmpList & RsSIN("ED_EMPNBR") & " - " & RsSIN("ED_SURNAME") & ", " & RsSIN("ED_FNAME") & vbNewLine
        RsSIN.MoveNext
    Loop
    fDupEmail_Act = xEmpList
        
    RsSIN.Close
    
    SQLQ = "SELECT ED_EMPNBR,ED_EMAIL, ED_SIN,ED_SSN,ED_SURNAME,ED_FNAME FROM TERM_HREMP "
    SQLQ = SQLQ & "WHERE ED_EMAIL = '" & xEmailAddr & "' "
    SQLQ = SQLQ & "And ED_EMPNBR <> " & xEmpnbr
    RsSIN.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    If Not RsSIN.EOF Then
        CheckDupEmpEmail = True
    End If
    
    xTerm_Emplist = ""
    If Not RsSIN.EOF Then
        xTerm_Emplist = "Terminated Employee(s):" & vbNewLine
    End If
    
    Do Until RsSIN.EOF
        xTerm_Emplist = xTerm_Emplist & RsSIN("ED_EMPNBR") & " - " & RsSIN("ED_SURNAME") & ", " & RsSIN("ED_FNAME") & vbNewLine
        RsSIN.MoveNext
    Loop
    fDupEmail_Term = xTerm_Emplist
        
    RsSIN.Close

    
End Function

' AC - dkostka - 05/08/2001 - Added function to find out if an ascii character is valid
' for a numeric entry field or not.
Private Function IsNumericEntry(KeyAscii As Integer, Optional NegAllowed As Boolean) As Boolean
    If KeyAscii = Asc(vbBack) Or IsNumeric(Chr(KeyAscii)) Or (NegAllowed And KeyAscii = Asc("-")) Then IsNumericEntry = True
End Function

Private Sub SwitchExpYearUserDefDate() 'Ticket #23428 Franks 03/20/2013
Dim I1 As Integer
Dim I2 As Integer
Dim K1 As Integer
Dim K2 As Integer
    'switch lable fields
    'Experience Year
    I1 = lblODate.Top ' Label3.Top
    K1 = lblLHire.Left ' Label3.Left
    'User Defined
    I2 = lblLHire.Top     'lblLHire top - lblUDay.Top
    K2 = lblODate.Left    'OMERS Date left - lblUDay.Left
    Label3.Top = I2
    Label3.Left = 3450 'K2
    Label3.Alignment = 1 'right
    lblUDay.Top = I1
    lblUDay.Left = K1
    lblUDay.Alignment = 0 'Left
    
    'switch textbox fields
    'Experience Year
    I1 = dlpDate(2).Top ' txtExpYear.Top
    K1 = dlpDate(5).Left ' txtExpYear.Left
    'User Defined
    I2 = dlpDate(5).Top     'dlpDate(3).Top
    K2 = dlpDate(2).Left    'dlpDate(3).Left
    txtExpYear.Top = I2
    txtExpYear.Left = K2 + 310
    dlpDate(3).Top = I1
    dlpDate(3).Left = K1  '- 310
    
End Sub

Private Sub OpenTranOutForm() 'Ticket #23903 Franks 06/19/2013
    If glbWFC And glbUserID = "999999999" Then
        MsgBox "The 999999999 account can not be used for Transfer Out.  Please log in under your employee number, and try again.", vbInformation + vbOKOnly, "Transfer Out Not Allowed"
        Exit Sub
    End If
    If Not gSec_Inq_Terminations Then
        MsgBox "You Do Not Have Authority For Employee Transfer Out"
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    Unload frmETERM
    glbTermTran = False
    Load frmETERM
    frmETERM.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub updFollowStatusToDate()
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset

On Error GoTo CrFollow_Err

'Ticket #28225 Franks 02/25/2016 WFC got record with 0 as employee # sometime here.
If glbLEE_ID = 0 Then Exit Sub

If OTDate = dlpDate(16).Text Then Exit Sub

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
If IsDate(OTDate) Then
    SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS='EDTO' "
    SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(OTDate)
ElseIf Not IsDate(OTDate) And IsDate(dlpDate(16).Text) Then
    SQLQ = SQLQ & " WHERE 1=2"
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If (Not IsDate(OTDate)) Or (rsTB.EOF And IsDate(dlpDate(16).Text)) Then
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
Else
    If Not IsDate(dlpDate(16).Text) Then
        If Not rsTB.EOF Then
            rsTB.Delete
            Msg = "A Follow Up Record was deleted!"
        Else
            'No follow up record to delete
        End If
        GoTo TheEnd
    Else
        Msg = "A Follow Up Record was updated!"
    End If
End If
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = glbLEE_ID
If IsDate(dlpDate(16).Text) Then 'Ticket #25142 Franks 03/06/2014
    rsTB("EF_FDATE") = CVDate(dlpDate(16).Text)
End If
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
End If
rsTB("EF_FREAS") = "EDTO"
rsTB("EF_COMMENTS") = "Employment Status To Date " & lStr("Follow-Up")
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update

Dim rsTT As New ADODB.Recordset
rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='EDTO'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsTT.EOF Then
    rsTT.AddNew
    rsTT("TB_COMPNO") = "001"
    rsTT("TB_NAME") = "FURE"
    rsTT("TB_KEY") = "EDTO"
    rsTT("TB_DESC") = "Employment Status To Date " & lStr("Follow-Up")
    rsTT("TB_LUSER") = glbUserID
    rsTT("TB_LDATE") = Date
    rsTT("TB_LTIME") = Time$
    rsTT.Update
End If
rsTT.Close

'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
'follow up record
Call Grant_FollowUpCode_Security(glbUserID, "EDTO", "Employment Status To Date " & lStr("Follow-Up"))

'Dim rsSR As New ADODB.Recordset
'rsSR.Open "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME='EDTO'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'If rsSR.EOF Then
'    'SQLQ = "INSERT INTO HR_SECURE_FOLLOW_UP(COMPNO,USERID," & Field_SQL("DESCRIPTION") & ",ACCESSABLE,Maintainable,CODENAME, TB_NAME) "
'    'SQLQ = SQLQ & " VALUES('001','" & glbSecUSERID & "'," & Chr$(34) & lStr(rsTD("TB_DESC")) & Chr$(34) & ",0,0,'" & rsTD("TB_KEY") & "','ECOM')"
'    rsSR.AddNew
'    rsSR("COMPNO") = "001"
'    rsSR("USERID") = glbUserID
'    rsSR("DESCRIPTION") = "Employment Status To Date " & lStr("Follow-Up")
'    rsSR("ACCESSABLE") = 1
'    rsSR("Maintainable") = 1
'    rsSR("CODENAME") = "EDTO"
'    rsSR("TB_NAME") = "FURE"
'   rsSR.Update
'End If
'rsSR.Close

TheEnd:
rsTB.Close
'MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP - Status To Date", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next
End Sub

Private Sub WFCNGSStartDate()
Dim xtemDate
    If glbEmpCountry = "U.S.A." Then
        If Not IsDate(dlpDate(24).Text) Then
            If clpCode(2).Text = "NONE" Or clpCode(2).Text = "EXEC" Then
                If Len(clpVadim1.Text) > 0 Then 'Ticket #25352 Franks 04/16/2014
                    If Len(getNGSSubGrpFromMatrix(rsDATA("ED_DIV"), clpCode(2).Text)) > 0 Then
                        dlpDate(24).Text = dlpDate(7).Text 'doh
                        dlpDate(29).Text = dlpDate(7).Text
                    End If
                End If
            Else
                'Ticket #20441 Franks 06/13/2011
                If Len(clpVadim1.Text) > 0 Then 'for Hourly NGS employees
                    If IsDate(dlpDate(7).Text) Then 'doh
                        xtemDate = DateAdd("D", 90, dlpDate(7).Text)
                        If NewHireForms.count > 0 Then 'New Hire only
                            If glbTrsHourWeek >= 20 Then 'Ticket #25248 Franks 03/24/2014
                                dlpDate(24).Text = xtemDate
                            End If
                        Else
                            'xtemDate = DateAdd("D", 91, dlpDate(7).Text)
                            'Msg = "Do you want to use the default NGS Start Date of " & xtemDate & "? "
                            'a% = MsgBox(Msg, 36, "Confirm ")
                            'If a% = 6 Then
                             If locWHRS > 0 Then 'Ticket #25248 Franks 03/24/2014
                                dlpDate(24).Text = xtemDate
                             End If
                            'End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub WFCUpdateBenefitEndDate(xEmpNo, xDATE, xType) 'Ticket #24179 Franks 02/25/2014
Dim rsBenT As New ADODB.Recordset
Dim SQLQ As String
    If Not IsDate(xDATE) Then
        Exit Sub
    End If
    If xType = "ALL" Then
        If glbMsgCustomVal = 4 Then 'remove Benefit Group
            clpBGroup.Text = ""
            SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = NULL WHERE ED_EMPNBR = " & xEmpNo & " "
            gdbAdoIhr001.Execute SQLQ
        End If
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = CVDate(xDATE)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
    If xType = "ComPaidNoIE" Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND NOT (BF_BCODE = 'IE') "
        SQLQ = SQLQ & "AND (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = CVDate(xDATE)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
    If xType = "RemoveEndDate" Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND not (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = Null ' CVDate(xDate)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
End Sub

Private Sub DispBenGrpFromUnion(xUnion) 'Frontenac  - Ticket #25122 Franks 03/07/2014
Dim rs As New ADODB.Recordset
Dim SQLQ As String
    If Len(xUnion) > 0 Then
        SQLQ = "SELECT * FROM HRMATRIX WHERE M_TYPE = 'UNIO' AND M_CODE = '" & xUnion & "' "
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("M_CONVERT1")) Then
                clpBGroup.Text = rs("M_CONVERT1")
            End If
        End If
        rs.Close
    End If
End Sub

Private Sub ShowHide_LOA_Attachment_Buttons()
    'Release 8.1
    'If LOA then show the option to view LOA Document
    If IsLOATypeCode(clpCode(1).Text) And IsDate(dlpDate(15).Text) And IsDate(dlpDate(16).Text) Then
        'Check if LOA Document is found
        If LOA_Document_Found(glbLEE_ID, clpCode(1).Text, dlpDate(15).Text, dlpDate(16).Text) <> 0 Then
            cmdImport2.Visible = True
            lblImport2.Visible = True
            imgNoSec2.Visible = False
            imgSec2.Visible = True
            cmdLOAComments.Visible = True
        Else
            cmdImport2.Visible = True
            lblImport2.Visible = True
            imgNoSec2.Visible = True
            imgSec2.Visible = False
            cmdLOAComments.Visible = True
        End If
    Else
        cmdImport2.Visible = False
        lblImport2.Visible = False
        imgNoSec2.Visible = False
        imgSec2.Visible = False
        cmdLOAComments.Visible = False
    End If

End Sub

Public Sub imgEmailBenefit_Click()
Dim xEmail
Dim xToEmail As String

On Error GoTo Email_Err
        
        'Release 8.1
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        
        'Ticket #30506 - Commenting the part where the email sending option is not given when Employee's Email is missing. The reason for doing so mainly is
        'because Employee's email is just being CC'd. The main email is going to email addresses specified on the Email Notification screen. And that's how
        'it has been done on Salary change - so maintaining the consistency.
        'If Len(xEmail) > 0 Then
            'Ticket #18090 - begin
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            Else
                'Ticket #20317 - More Emails for everyone
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            End If
            'Ticket #18090 - end
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONBENEFIT")
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
            Else
                frmSendEmail.txtCC.Text = xEmail
            End If
            'frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice"
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR Benefit Update Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
        '    If Len(glbLEE_SName) = 0 Then
        '        MsgBox "There is no email on Status/Dates screen for employee. "
        '    Else
        '        MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
        '    End If
        'End If

    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Private Sub WFCUScreenSetup() 'Ticket #28515 Franks 04/26/2016
    lblRegion.Left = 6720
    lblRegion.Top = 480
    lblRegion.Caption = lStr("Region")
    lblRegion.FontBold = True
    lblRegion.Visible = True
    clpCode(7).Left = 8070
    clpCode(7).Top = 430
    clpCode(7).DataField = "ED_REGION"
    clpCode(7).Visible = True
End Sub

Private Sub WFCSetUnionDate() 'Ticket #30376 Franks 07/17/2017
'On new hire, if UNION is entered, default the effective date to DOH
If NewHireForms.count > 0 Then
    If Len(clpCode(2).Text) > 0 Then
        If clpCode(2).Caption = "Unassigned" Then
        Else
            If Len(dlpDate(36).Text) = 0 Then
                dlpDate(36).Text = dlpDate(7).Text ' DOH
            End If
        End If
    End If
End If
End Sub
