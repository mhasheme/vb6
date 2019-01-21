VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEMERG 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Emergency Contact "
   ClientHeight    =   8355
   ClientLeft      =   135
   ClientTop       =   2535
   ClientWidth     =   11235
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8355
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   2160
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   9
      Top             =   7815
      Visible         =   0   'False
      Width           =   11235
      _Version        =   65536
      _ExtentX        =   19817
      _ExtentY        =   952
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7140
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
   Begin INFOHR_Controls.DateLookup dlpEXPIRYDATE 
      DataField       =   "ED_EXPIRYDATE"
      Height          =   285
      Left            =   6210
      TabIndex        =   35
      Tag             =   "40-Expiry Date"
      Top             =   5400
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Frame frmDetail 
      Caption         =   "PRIMARY DOCTOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Top             =   3480
      Width           =   4635
      Begin VB.TextBox txtDORADDRESS 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDORADDRESS2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1650
         MaxLength       =   60
         TabIndex        =   29
         Tag             =   "00-Primary Doctor-Address 2"
         Top             =   1400
         Width           =   2775
      End
      Begin VB.TextBox txtDORADDRESS 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDORADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1650
         MaxLength       =   60
         TabIndex        =   28
         Tag             =   "00-Primary Doctor- Address 1"
         Top             =   1040
         Width           =   2775
      End
      Begin VB.TextBox txtEDoctor 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDOCTOR"
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
         Left            =   1650
         MaxLength       =   60
         TabIndex        =   26
         Tag             =   "00-Primary Doctor-Doctor Name"
         Top             =   360
         Width           =   2775
      End
      Begin MSMask.MaskEdBox medDTele 
         DataField       =   "ED_EDPNBR"
         Height          =   285
         Index           =   0
         Left            =   1650
         TabIndex        =   27
         Tag             =   "10-Primary Doctor-Telephone"
         Top             =   690
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
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
         Left            =   90
         TabIndex        =   59
         Top             =   1395
         Width           =   705
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1"
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
         Left            =   90
         TabIndex        =   58
         Top             =   1065
         Width           =   705
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
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
         Left            =   90
         TabIndex        =   57
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone "
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
         Left            =   90
         TabIndex        =   56
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame frmDetail 
      Caption         =   "SECONDARY DOCTOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   4920
      TabIndex        =   52
      Top             =   3480
      Width           =   4575
      Begin VB.TextBox txtDOR2ADDRESS 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDOR2ADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1590
         MaxLength       =   60
         TabIndex        =   32
         Tag             =   "00-Secondary Doctor- Address 1"
         Top             =   1040
         Width           =   2805
      End
      Begin VB.TextBox txtDOR2ADDRESS 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDOR2ADDRESS2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1590
         MaxLength       =   60
         TabIndex        =   33
         Tag             =   "00-Secondary Doctor-Address 2"
         Top             =   1400
         Width           =   2805
      End
      Begin VB.TextBox txtEDoctor 
         Appearance      =   0  'Flat
         DataField       =   "ED_EDOCTOR2"
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
         Left            =   1590
         MaxLength       =   60
         TabIndex        =   30
         Tag             =   "00-Secondary Doctor-Doctor Name"
         Top             =   360
         Width           =   2805
      End
      Begin MSMask.MaskEdBox medDTele 
         DataField       =   "ED_EDPNBR2"
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   31
         Tag             =   "10-Secondary Doctor-Telephone"
         Top             =   690
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1"
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
         Left            =   120
         TabIndex        =   61
         Top             =   1065
         Width           =   705
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
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
         Left            =   120
         TabIndex        =   60
         Top             =   1395
         Width           =   705
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone "
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
         Left            =   120
         TabIndex        =   54
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
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
         TabIndex        =   53
         Top             =   405
         Width           =   1215
      End
   End
   Begin VB.Frame frmDetail 
      Caption         =   "SECONDARY CONTACT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Index           =   2
      Left            =   4920
      TabIndex        =   43
      Top             =   480
      Width           =   4575
      Begin VB.TextBox txtCRelate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_RELATE2"
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
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox comRelation 
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
         Index           =   1
         Left            =   1620
         TabIndex        =   20
         Tag             =   "00-Secondary Contact-Relationship to Employee"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtContName 
         Appearance      =   0  'Flat
         DataField       =   "ED_ECONT2"
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
         MaxLength       =   25
         TabIndex        =   19
         Tag             =   "00-Secondary Contact-Contact Name"
         Top             =   390
         Width           =   2775
      End
      Begin VB.TextBox txtEEMail 
         Appearance      =   0  'Flat
         DataField       =   "ED_EEMAIL2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1620
         MaxLength       =   60
         TabIndex        =   25
         Tag             =   "00-Secondary Contact-Email Address"
         Top             =   2490
         Width           =   2775
      End
      Begin MSMask.MaskEdBox medCTele 
         DataField       =   "ED_ENBR2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   21
         Tag             =   "10-Secondary Contact-Telephone"
         Top             =   1110
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "(###) ###-####    Ext(######)"
         Mask            =   "(###) ###-####    Ext(######)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCTele 
         DataField       =   "ED_EP2NBR2"
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   22
         Tag             =   "10-Secondary Contact-Alternate Telephone"
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "(###) ###-####    Ext(######)"
         Mask            =   "(###) ###-####    Ext(######)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medECellPhone 
         DataField       =   "ED_ECELLPHONE2"
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   23
         Tag             =   "10-Secondary Contact-Cellular Telephone Number"
         Top             =   1770
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox medEPageNbr 
         DataField       =   "ED_EPAGENBR2"
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   24
         Tag             =   "10-Secondary Contact-Pager Number"
         Top             =   2130
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
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
         Left            =   90
         TabIndex        =   51
         Top             =   750
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone "
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
         Left            =   90
         TabIndex        =   50
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
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
         Left            =   90
         TabIndex        =   49
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone #2"
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
         Left            =   90
         TabIndex        =   48
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pager Number"
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
         Left            =   90
         TabIndex        =   47
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cellular Telephone"
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
         Left            =   90
         TabIndex        =   46
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label lblTitle 
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
         Index           =   16
         Left            =   90
         TabIndex        =   45
         Top             =   2490
         Width           =   1200
      End
   End
   Begin VB.Frame frmDetail 
      Caption         =   "PRIMARY CONTACT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4635
      Begin VB.TextBox txtEEMail 
         Appearance      =   0  'Flat
         DataField       =   "ED_EEMAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1620
         MaxLength       =   60
         TabIndex        =   18
         Tag             =   "00-Primary Contact-Email Address"
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtCRelate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_RELATE"
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
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox comRelation 
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
         Index           =   0
         ItemData        =   "feemerg.frx":0000
         Left            =   1620
         List            =   "feemerg.frx":0002
         TabIndex        =   12
         Tag             =   "01-Primary Contact-Relationship to Employee"
         Top             =   690
         Width           =   2055
      End
      Begin VB.TextBox txtContName 
         Appearance      =   0  'Flat
         DataField       =   "ED_ECONT"
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
         MaxLength       =   25
         TabIndex        =   11
         Tag             =   "01-Primary Contact-Contact Name"
         Top             =   360
         Width           =   2775
      End
      Begin MSMask.MaskEdBox medCTele 
         DataField       =   "ED_ENBR"
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   14
         Tag             =   "10-Primary Contact-Telephone"
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "(###) ###-####    Ext(######)"
         Mask            =   "(###) ###-####    Ext(######)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCTele 
         DataField       =   "ED_EP2NBR"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   15
         Tag             =   "10-Primary Contact-Alternate Telephone"
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "(###) ###-####    Ext(######)"
         Mask            =   "(###) ###-####    Ext(######)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medECellPhone 
         DataField       =   "ED_ECELLPHONE"
         DataSource      =   "Data1"
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   16
         Tag             =   "10-Primary Contact-Cellular Telephone Number"
         Top             =   1770
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox medEPageNbr 
         DataField       =   "ED_EPAGENBR"
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   17
         Tag             =   "10-Primary Contact-Pager Number"
         Top             =   2130
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
      Begin VB.Label lblTitle 
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
         Index           =   15
         Left            =   90
         TabIndex        =   65
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
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
         Left            =   90
         TabIndex        =   42
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone "
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
         Left            =   90
         TabIndex        =   41
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
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
         Left            =   90
         TabIndex        =   40
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone #2"
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
         Left            =   90
         TabIndex        =   39
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pager Number"
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
         Left            =   90
         TabIndex        =   38
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cellular Telephone"
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
         Left            =   90
         TabIndex        =   37
         Top             =   1830
         Width           =   1320
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ED_LUSER"
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
      Left            =   7320
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ED_LTIME"
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
      Left            =   5760
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
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
      Left            =   4200
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11235
      _Version        =   65536
      _ExtentX        =   19817
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
         Left            =   7320
         MaxLength       =   25
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   90
         Visible         =   0   'False
         Width           =   1230
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
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   90
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
         Left            =   8040
         TabIndex        =   67
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblEEnum 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Height          =   240
         Left            =   5040
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   3150
         TabIndex        =   3
         Top             =   120
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medVERSION 
      DataField       =   "ED_VERSION"
      Height          =   285
      Left            =   1770
      TabIndex        =   36
      Tag             =   "00-Version #"
      Top             =   5730
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   6
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
   Begin MSMask.MaskEdBox medHEALTHCARD 
      DataField       =   "ED_HEALTHCARD"
      Height          =   285
      Left            =   1770
      TabIndex        =   34
      Tag             =   "00-Health Card #"
      Top             =   5400
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date"
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
      Left            =   5010
      TabIndex        =   66
      Top             =   5445
      Width           =   810
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version #"
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
      Left            =   210
      TabIndex        =   64
      Top             =   5760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Health Card #"
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
      Left            =   210
      TabIndex        =   63
      Top             =   5460
      Width           =   990
   End
End
Attribute VB_Name = "frmEMERG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim fglbNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim oContName, oContNbr1, oContNbr2, oEmail
Dim oConRelation1

Private Sub AUDITEMERG()
Dim xBatchID
Dim HRChanges As New Collection
'Aurora and Kawartha Lakes and City of Niagara Falls
'Ticket #24565 - Do not transfer for District Municipality of Muskoka
If Not glbCompSerial = "S/N - 2378W" And Not glbCompSerial = "S/N - 2363W" And Not glbCompSerial = "S/N - 2276W" And _
    Not glbCompSerial = "S/N - 2373W" Then
    Call isChanged_Field(HRChanges, oContName, txtContName(0))
    Call isChanged_Field(HRChanges, oContNbr1, medCTele(0))
    Call isChanged_Field(HRChanges, oContNbr2, medCTele(2))
    Call isChanged_Field(HRChanges, oEmail, txtEEMail(0))
    Call isChanged_Field(HRChanges, oConRelation1, comRelation(0))
    Call Passing_Changes(HRChanges, Contacts, "M", Date, glbLEE_ID)
End If

'Ticket #22682 - Release 8.0: To show the Emergency Contact Audit - but can only show if the Audit table is updated for
'everyone
'If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
Dim rsAU2 As New ADODB.Recordset
Dim rsAU As New ADODB.Recordset
Dim UpdateAudit2 As Boolean
Dim rsEmp2 As New ADODB.Recordset
Dim xDiv As String
    UpdateAudit2 = False
    'Ticket #25017 Franks - worked with Hemu, moved the following to above
    'If isChanged_Field(HRChanges, oContName, txtContName(0)) Then UpdateAudit2 = True
    'If isChanged_Field(HRChanges, oContNbr1, medCTele(0)) Then UpdateAudit2 = True
    'If isChanged_Field(HRChanges, oConRelation1, comRelation(0)) Then UpdateAudit2 = True
    
    If UpdateAudit2 Then
        'Ticket #12843
        rsEmp2.Open "SELECT ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " ", gdbAdoIhr001, adOpenStatic
        xDiv = ""
        If Not rsEmp2.EOF Then
            If Not IsNull(rsEmp2("ED_DIV")) Then xDiv = rsEmp2("ED_DIV")
        End If
        rsEmp2.Close
        
        'Ticket #22682 - Release 8.0: Had to add blank record in HRAUDIT for the Audit report to work
        rsAU.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        rsAU.AddNew
        rsAU("AU_NEWEMP") = "N"
        rsAU("AU_TYPE") = "M"
        rsAU("AU_COMPNO") = "001"
        rsAU("AU_EMPNBR") = glbLEE_ID
        rsAU("AU_LDATE") = Date
        rsAU("AU_LUSER") = glbUserID
        rsAU("AU_LTIME") = Time$
        rsAU("AU_UPLOAD") = "N"
        rsAU.Update
        If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
            Call FamilyDayAuditSync(glbLEE_ID, rsAU)
        End If
        rsAU.Close
        
        rsAU2.Open "SELECT * FROM HRAUDIT2 WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        rsAU2.AddNew
        rsAU2("AU_NEWEMP") = "N"
        rsAU2("AU_TYPE") = "M"
        rsAU2("AU_COMPNO") = "001"
        rsAU2("AU_EMPNBR") = glbLEE_ID
        rsAU2("AU_LDATE") = Date
        rsAU2("AU_LUSER") = glbUserID
        rsAU2("AU_LTIME") = Time$
        rsAU2("AU_UPLOAD") = "N"
        If oContName <> txtContName(0) Then
            rsAU2("AU_ECONT") = txtContName(0)
        End If
        If oContNbr1 <> medCTele(0).Text Then
            rsAU2("AU_ENBR") = medCTele(0)
        End If
        If oConRelation1 <> comRelation(0) Then
            rsAU2("AU_RELATE") = comRelation(0)
        End If
        If Len(xDiv) > 0 Then
            rsAU2("AU_DIVUPL") = xDiv
        End If
        rsAU2.Update
    End If
'End If

End Sub

Public Sub cmdCancel_Click()
Dim x
'''On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack   '23June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEMERG" Then glbOnTop = ""

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Public Sub cmdModify_Click()

'''On Error GoTo Mod_Err


oContName = txtContName(0)
oContNbr1 = medCTele(0)
oContNbr2 = medCTele(2)
oEmail = txtEEMail(0)
oConRelation1 = comRelation(0)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREMP", "Modify")
Call RollBack  '23June99 - js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()

'''On Error GoTo Add_Err

If Not chkEmerg() Then Exit Sub '17Aug99 js
Call AUDITEMERG
rsDATA.Requery
Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

rsDATA.Update
fglbNew = False
Call Employee_Master_Integration(glbLEE_ID)
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Call EERetrieve

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    Call FamilyDayEmpSync(glbLEE_ID)
End If

Call NextForm
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack  '23June99 - js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Emegency Contact Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Emegency Contact Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'Call setRptLabel(Me, 1)
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "Rgcontct.rpt"
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
    xReport = glbIHRREPORTS & "Rgcontc2.rpt"
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

RHeading = lblEEName & "'s Emegency Contact Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Emegency Contact Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'Call setRptLabel(Me, 1)
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "Rgcontct.rpt"
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
    xReport = glbIHRREPORTS & "Rgcontc2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub comRelation_Change(Index As Integer)
comRelation_Click (Index)
End Sub

Private Sub comRelation_Click(Index As Integer)
Dim tlen As Integer

tlen = Len(comRelation(Index).Text)

If tlen > 20 Then tlen = 20
If tlen >= 1 Then
    txtCRelate(Index).Text = Left$(comRelation(Index).Text, tlen)
Else
    txtCRelate(Index).Text = ""
End If

End Sub

Private Sub comRelation_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

'''On Error GoTo EERError

Screen.MousePointer = HOURGLASS

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

If rsDATA.BOF And rsDATA.EOF Then
   MsgBox "Sorry, Employee Removed prior to your access"
Else
   EERetrieve = True
End If



Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack  '23June99 - js

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMEMERG"
    Call SetPhone
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEMERG"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEMERG"

Screen.MousePointer = HOURGLASS
If glbtermopen Then
Data1.ConnectionString = glbAdoIHRAUDIT
Else
Data1.ConnectionString = glbAdoIHRDB
End If

Call LdcomRel(0)
Call LdcomRel(1)
Screen.MousePointer = DEFAULT
'Frank - Mar 04,04 Begin - Ticket #5733
'Hide Health Card information as Jerry's request for v7.2
'Changed the three fields of Health Card to invisible
'Frank - Mar 04,04 End

'Hemu - 11/17/2003 Begin  - Ticket # 5104
'If glbCompSerial = "S/N - 2234W" Then  'laura 03/05/98
'    lblTitle(5).Caption = "Business #"
'    lblTitle(4).Caption = "Health Card #"
'    lblTitle(2).Visible = False
'    lblTitle(6).Caption = "Residence #"
'    lblTitle(8).Visible = False
'    comRelation(1).Visible = False
'End If
'Hemu - 11/17/2003 End


If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

Call SetPhone

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call ST_UPD_MODE(False)
If Not gSec_Upd_EmergContacts Then
'    cmdModify.Enabled = False
End If
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'Hemu - 04/28/2004 - Begin - Health Card Info. display - Serial # control
If glbCompSerial = "S/N - 2225W" Or glbCompSerial = "S/N - 2190W" Or glbLinamar Then
    'lblTitle(23).Visible = True
    lblTitle(24).Visible = True
    'lblTitle(25).Visible = True
    'medHEALTHCARD.Visible = True
    medVERSION.Visible = True
    'dlpEXPIRYDATE.Visible = True
    lblTitle(23).WordWrap = False
    lblTitle(23).Caption = "Health Card #"
    medHEALTHCARD.Tag = "00-Health Card #"
Else
    'As per Next Release Documentation - Show Health fields
    lblTitle(23).WordWrap = True
    lblTitle(23).Caption = "Health/Insurance Card #"
    medHEALTHCARD.Tag = "00-Health/Insurance Card #"
End If
'Hemu - 04/28/2004 - End

If glbCompSerial = "S/N - 2394W" Then  'St. John Ticket #15279
    lblTitle(18).Caption = "City, Prov && PC"
    lblTitle(19).Caption = "City, Prov && PC"
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    lblTitle(3).FontBold = True
    lblTitle(5).FontBold = True
    lblTitle(7).FontBold = True
End If

'Ticket #27818 - County of Wellington - Grey out Physician and Health card info.
If glbCompSerial = "S/N - 2262W" Then
    frmDetail(1).Enabled = False
    frmDetail(3).Enabled = False
    medHEALTHCARD.Enabled = False
    medVERSION.Enabled = False
    dlpEXPIRYDATE.Enabled = False
End If

End Sub
Private Sub SetPhone()

    If Not (glbEmpCountry = "CANADA" Or glbEmpCountry = "U.S.A." Or glbEmpCountry = "MEXICO") Then
        medCTele(0).MaxLength = 25
        medCTele(0).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medCTele(1).MaxLength = 25
        medCTele(1).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medCTele(2).MaxLength = 25
        medCTele(2).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medCTele(3).MaxLength = 25
        medCTele(3).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medECellPhone(0).MaxLength = 25
        medECellPhone(0).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medECellPhone(1).MaxLength = 25
        medECellPhone(1).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medEPageNbr(0).MaxLength = 25
        medEPageNbr(0).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medEPageNbr(1).MaxLength = 25
        medEPageNbr(1).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medDTele(0).MaxLength = 25
        medDTele(0).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medDTele(1).MaxLength = 25
        medDTele(1).Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
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
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
End Sub

Public Sub imgEmail_Click()
    Call txtEEMail_DblClick(0)
End Sub

Private Sub lblEEID_Change()

If Len(txtSurname.Text) > 0 And Len(txtFName.Text) > 0 Then  ' dont do on add new until in
    frmEMERG.Caption = "Emergency Contacts - " & Left$(txtSurname, 5)
    frmEMERG.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If
lblEEnum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub LdcomRel(IDX%)

comRelation(IDX).AddItem "Aunt"
comRelation(IDX).AddItem "Brother"
comRelation(IDX).AddItem "Children"
comRelation(IDX).AddItem "Daughter"
comRelation(IDX).AddItem "Doctor"
comRelation(IDX).AddItem "Domestic"
comRelation(IDX).AddItem "Ex-Spouse"
comRelation(IDX).AddItem "Father"
comRelation(IDX).AddItem "Fiance"
comRelation(IDX).AddItem "Friend"
comRelation(IDX).AddItem "Husband"
comRelation(IDX).AddItem "Mother"
comRelation(IDX).AddItem "Parents"
comRelation(IDX).AddItem "Relative"
comRelation(IDX).AddItem "Sister"
comRelation(IDX).AddItem "Son"
comRelation(IDX).AddItem "Spouse"
comRelation(IDX).AddItem "Uncle"
comRelation(IDX).AddItem "Wife"
comRelation(IDX).AddItem "Mother in-law"
comRelation(IDX).AddItem "Father in-law"
comRelation(IDX).AddItem "Brother in-law"
comRelation(IDX).AddItem "Sister in-law"

End Sub



Private Sub medCTele_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCTele_ValidationError(Index As Integer, InvalidText As String, StartPosition As Integer)
    glbBYPASS380 = True
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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
  
comRelation(0).Enabled = TF
comRelation(1).Enabled = TF
medCTele(0).Enabled = TF
medCTele(1).Enabled = TF
medCTele(2).Enabled = TF
medCTele(3).Enabled = TF
txtContName(0).Enabled = TF
txtContName(1).Enabled = TF
txtEDoctor(0).Enabled = TF
txtEDoctor(1).Enabled = TF
txtEEMail(0).Enabled = TF
txtEEMail(1).Enabled = TF
medECellPhone(0).Enabled = TF
medECellPhone(1).Enabled = TF
medEPageNbr(0).Enabled = TF
medEPageNbr(1).Enabled = TF
medDTele(0).Enabled = TF
medDTele(1).Enabled = TF
txtDORADDRESS(0).Enabled = TF
txtDORADDRESS(1).Enabled = TF
txtDOR2ADDRESS(0).Enabled = TF
txtDOR2ADDRESS(1).Enabled = TF
'Ticket #27818 - County of Wellington - Grey out Physician and Health card info.
If glbCompSerial <> "S/N - 2262W" Then
    medHEALTHCARD.Enabled = TF
    medVERSION.Enabled = TF
    dlpEXPIRYDATE.Enabled = TF
End If
End Sub

Private Sub medDTele_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medECellPhone_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEPageNbr_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHEALTHCARD_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medVERSION_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtContName_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtCRelate_Change(Index As Integer)
    comRelation(Index).Text = txtCRelate(Index).Text
End Sub
                   
Function chkEmerg() '------17Aug99 js-----------
Dim x%

'''On Error GoTo chkEmerg_Err

chkEmerg = False

'Hemu - 18/08/2003 Begin - Jerry asked to remove the required field feature
'If Len(txtContName(0).Text) = 0 Then
'    MsgBox "Primary contact name is a required field."
'    txtContName(0).SetFocus
'    Exit Function
'End If
'Hemu - 18/08/2003 End

'If Len(medCTele(0)) = 0 Then
'    MsgBox "Telephone number is a required field."
'    medCTele(0).SetFocus
'    Exit Function
'End If

'Hemu - 18/08/2003 Begin - Jerry asked to remove the required field feature
'If Len(txtCRelate(0)) <= 0 Then
'    MsgBox "Relationship is a required field."
'    comRelation(0).SetFocus
'    Exit Function
'End If
'Hemu - 18/08/2003 End

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    If Len(txtContName(0).Text) = 0 Then
        MsgBox "Primary contact name is a required field."
        txtContName(0).SetFocus
        Exit Function
    End If
    If Len(txtCRelate(0)) <= 0 Then
        MsgBox "Relationship is a required field."
        comRelation(0).SetFocus
        Exit Function
    End If
    If Len(medCTele(0)) = 0 Then
        MsgBox "Telephone number is a required field."
        medCTele(0).SetFocus
        Exit Function
    End If
End If

If Len(dlpEXPIRYDATE) > 0 Then
    If Not IsDate(dlpEXPIRYDATE) Then
        MsgBox "Invalid Expiry Date"
        dlpEXPIRYDATE.SetFocus
        Exit Function
    End If
End If

chkEmerg = True

Exit Function

chkEmerg_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEmerg", "HREMP", "edit/Add")
Call RollBack '17Aug99 js

End Function

Private Function RollBack()
'''On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function


Private Sub txtDOR2ADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDORADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEDoctor_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

'Hemu - 07/22/2003 Begin - Sending email
Private Sub txtEEMail_DblClick(Index As Integer)
'''On Error GoTo Email_Err
    If gsEMAIL_SENDING Then
        If Len(txtEEMail(Index).Text) > 0 Then
            frmSendEmail.txtTo.Text = txtEEMail(Index).Text
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
'Hemu - 07/22/2003 End

Private Sub txtEEMail_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub
Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_ECONT, "
SQLQ = SQLQ & "ED_RELATE, ED_ENBR, ED_EP2NBR, ED_ECELLPHONE,"
SQLQ = SQLQ & "ED_EPAGENBR, ED_EEMAIL, ED_EDOCTOR, ED_EDPNBR,"
SQLQ = SQLQ & "ED_ECONT2, ED_RELATE2, ED_ENBR2, ED_EP2NBR2,"
SQLQ = SQLQ & "ED_ECELLPHONE2, ED_EPAGENBR2, ED_EEMAIL2,"
SQLQ = SQLQ & "ED_EDOCTOR2, ED_EDPNBR2, ED_EDORADDRESS,"
SQLQ = SQLQ & "ED_EDORADDRESS2, ED_EDOR2ADDRESS,"

SQLQ = SQLQ & "ED_HEALTHCARD, ED_VERSION, ED_EXPIRYDATE,"

SQLQ = SQLQ & "ED_EDOR2ADDRESS2, ED_LDATE, ED_LTIME,"

SQLQ = SQLQ & "ED_LUSER"
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ
End Function

Public Sub Display_Value()
    Dim SQLQ
    If rsDATA.EOF Or rsDATA.BOF Then
        Call Set_Control("B", Me)
        Call SET_UP_MODE
        Exit Sub
    End If
    
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
'Data1.RecordSource = SQLQ
'Data1.Refresh
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
    Me.cmdModify_Click
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
UpdateRight = gSec_Upd_EmergContacts
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


