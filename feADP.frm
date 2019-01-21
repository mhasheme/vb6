VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmpADP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee ADP Information"
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
   Begin VB.Frame frmUserFields 
      Caption         =   "User Defined Fields"
      Height          =   2535
      Left            =   240
      TabIndex        =   32
      Top             =   3720
      Width           =   8415
      Begin VB.ComboBox cmbStatusFlag5 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Tag             =   "User Defined Field 5"
         Top             =   1905
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cmbStatusFlag4 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "User Defined Field 4"
         Top             =   1545
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cmbStatusFlag3 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "User Defined Field 3"
         Top             =   1185
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cmbStatusFlag2 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Tag             =   "User Defined Field 2"
         Top             =   825
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cmbStatusFlag1 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Tag             =   "User Defined Field 1"
         Top             =   465
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox txtUValue 
         Appearance      =   0  'Flat
         DataField       =   "AP_TXT5"
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
         Index           =   4
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   12
         Tag             =   "00-Field Value 5"
         Top             =   1920
         Width           =   4680
      End
      Begin VB.TextBox txtUValue 
         Appearance      =   0  'Flat
         DataField       =   "AP_TXT4"
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
         Index           =   3
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "00-Field Value 4"
         Top             =   1560
         Width           =   4680
      End
      Begin VB.TextBox txtUValue 
         Appearance      =   0  'Flat
         DataField       =   "AP_TXT3"
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
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "00-Field Value 3"
         Top             =   1200
         Width           =   4680
      End
      Begin VB.TextBox txtUValue 
         Appearance      =   0  'Flat
         DataField       =   "AP_TXT2"
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
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "00-Field Value 2"
         Top             =   840
         Width           =   4680
      End
      Begin VB.TextBox txtUValue 
         Appearance      =   0  'Flat
         DataField       =   "AP_TXT1"
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
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "00-Field Value 1"
         Top             =   480
         Width           =   4680
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "User Defined Field 5"
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
         Index           =   14
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "User Defined Field 4"
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
         Index           =   13
         Left            =   240
         TabIndex        =   36
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "User Defined Field 2"
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
         Index           =   11
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "User Defined Field 1"
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
         Index           =   10
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "User Defined Field 3"
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
         Index           =   12
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1440
      End
   End
   Begin VB.Frame frmADPFields 
      Caption         =   "ADP Fields"
      Height          =   2775
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   8415
      Begin VB.TextBox txtCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_CLOCKP2"
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
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   42
         Tag             =   "00-Clock Position 2"
         Top             =   1038
         Width           =   240
      End
      Begin VB.ComboBox cmbDataCtrlFull 
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "Data Control Full"
         Top             =   240
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cmbClockFull 
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "Clock Full"
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtDCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_DCP1"
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
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   4
         Tag             =   "00-Data Control Position 1"
         Top             =   650
         Width           =   240
      End
      Begin VB.TextBox txtDCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_DCP2"
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
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "00-Data Control Position 2"
         Top             =   1040
         Width           =   240
      End
      Begin VB.TextBox txtDCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_DCP3"
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
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "00-Data Control Position 3"
         Top             =   1430
         Width           =   240
      End
      Begin VB.TextBox txtDCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_DCP4"
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
         Index           =   3
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "00-Data Control Position 4"
         Top             =   1820
         Width           =   240
      End
      Begin VB.TextBox txtCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_CLOCKP1"
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
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   0
         Tag             =   "00-Clock Position 1"
         Top             =   650
         Width           =   240
      End
      Begin VB.TextBox txtCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_CLOCKP3"
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
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "00-Clock Position 3"
         Top             =   1426
         Width           =   240
      End
      Begin VB.TextBox txtCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_CLOCKP4"
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
         Index           =   3
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "00-Clock Position 4"
         Top             =   1814
         Width           =   240
      End
      Begin VB.TextBox txtCP 
         Appearance      =   0  'Flat
         DataField       =   "AP_CLOCKP5"
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
         Index           =   4
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "00-Clock Position 5"
         Top             =   2205
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data Control Position 2"
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
         Index           =   7
         Left            =   3960
         TabIndex        =   43
         Top             =   1085
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data Control Full"
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
         Index           =   16
         Left            =   3960
         TabIndex        =   41
         Top             =   300
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Clock Full"
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
         Index           =   15
         Left            =   240
         TabIndex        =   40
         Top             =   300
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data Control Position 1"
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
         Index           =   5
         Left            =   3960
         TabIndex        =   31
         Top             =   695
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data Control Position 3"
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
         Index           =   8
         Left            =   3960
         TabIndex        =   30
         Top             =   1475
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Data Control Position 4"
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
         Index           =   9
         Left            =   3960
         TabIndex        =   29
         Top             =   1865
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clock Position 2"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1083
         Width           =   2100
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clock Position 3"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1471
         Width           =   2100
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clock Position 5"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2250
         Width           =   1740
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Clock Position 1"
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
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   695
         Width           =   2340
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clock Position 4"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1859
         Width           =   1740
      End
   End
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
      TabIndex        =   20
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AP_LUSER"
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
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AP_LTIME"
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
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AP_LDATE"
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
      TabIndex        =   17
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
      TabIndex        =   16
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
         Left            =   7320
         TabIndex        =   49
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
         TabIndex        =   21
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
         TabIndex        =   13
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "AP_EMPNBR"
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
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   120
         Width           =   1740
      End
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "AP_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3000
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmEmpADP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim fglbNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim oContName, oContNbr1, oContNbr2, oEmail
Dim oCP(4), oDC(3), oUFN(4), oUFV(4)
Dim xAction As String

Public Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

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
If glbOnTop = "FRMEMPADP" Then glbOnTop = ""

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Public Sub cmdModify_Click()

On Error GoTo Mod_Err
Dim X As Integer

For X = 0 To 4
    oCP(X) = txtCP(X)
    If X <> 4 Then oDC(X) = txtDCP(X)
    'oUFN(x) = txtUName(x)
    oUFV(X) = txtUValue(X)
Next X
xAction = "M"
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

On Error GoTo Add_Err

If Not chkEmpOther() Then Exit Sub '17Aug99 js

'rsDATA.Requery
Call UpdUStats(Me) ' update user's stats (who did it and when)
Call AUDITBENF(xAction, "1")
Call Set_Control("U", Me, rsDATA)
If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Call EERetrieve

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
Dim RHeading As String, xReport, X%

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
        For X% = 0 To 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
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
Dim RHeading As String, xReport, X%

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
        For X% = 0 To 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
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

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

Call ADP_DropDown_Controls

If glbtermopen Then
    SQLQ = "Select * from Term_HR_ADP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select * from HR_ADP "
    SQLQ = SQLQ & " where AP_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Then
    rsDATA.AddNew
    rsDATA("AP_COMPNO") = "001"
    rsDATA("AP_EMPNBR") = glbLEE_ID
    If glbtermopen Then
        rsDATA("TERM_SEQ") = glbTERM_Seq
    End If
    rsDATA.Update
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Call Display_Value

'If rsDATA.BOF And rsDATA.EOF Then
'   MsgBox "Sorry, Employee Removed prior to your access"
'Else
   EERetrieve = True
'End If



Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack  '23June99 - js

Exit Function

End Function

Private Sub dlpPassDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpPermitDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbClockFull_Change()
    If cmbClockFull.ListIndex > -1 Then
        Call Emp_ADP_ClockFull
    End If
End Sub

Private Sub cmbClockFull_Click()
    If cmbClockFull.ListIndex > -1 Then
        Call Emp_ADP_ClockFull
    End If
End Sub

Private Sub cmbClockFull_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbDataCtrlFull_Change()
    If cmbDataCtrlFull.ListIndex > -1 Then
        Call Emp_ADP_DataCtrlFull
    End If
End Sub

Private Sub cmbDataCtrlFull_Click()
    If cmbDataCtrlFull.ListIndex > -1 Then
        Call Emp_ADP_DataCtrlFull
    End If
End Sub

Private Sub cmbDataCtrlFull_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbStatusFlag1_Change()
    If cmbStatusFlag1.ListIndex > -1 Then
        txtUValue(0).Text = Left(cmbStatusFlag1.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag1_Click()
    If cmbStatusFlag1.ListIndex > -1 Then
        txtUValue(0).Text = Left(cmbStatusFlag1.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag1_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbStatusFlag2_Change()
    If cmbStatusFlag2.ListIndex > -1 Then
        txtUValue(1).Text = Left(cmbStatusFlag2.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag2_Click()
    If cmbStatusFlag2.ListIndex > -1 Then
        txtUValue(1).Text = Left(cmbStatusFlag2.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag2_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbStatusFlag3_Change()
    If cmbStatusFlag3.ListIndex > -1 Then
        txtUValue(2).Text = Left(cmbStatusFlag3.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag3_Click()
    If cmbStatusFlag3.ListIndex > -1 Then
        txtUValue(2).Text = Left(cmbStatusFlag3.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag3_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbStatusFlag4_Change()
    If cmbStatusFlag4.ListIndex > -1 Then
        txtUValue(3).Text = Left(cmbStatusFlag4.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag4_Click()
    If cmbStatusFlag4.ListIndex > -1 Then
        txtUValue(3).Text = Left(cmbStatusFlag4.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag4_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbStatusFlag5_Change()
    If cmbStatusFlag5.ListIndex > -1 Then
        txtUValue(4).Text = Left(cmbStatusFlag5.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag5_Click()
    If cmbStatusFlag5.ListIndex > -1 Then
        txtUValue(4).Text = Left(cmbStatusFlag5.Text, 1)
    End If
End Sub

Private Sub cmbStatusFlag5_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMEMPADP"
    Call SET_UP_MODE
    txtCP(0).SetFocus
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEMPADP"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim i
Dim xDiv
glbOnTop = "FRMEMPADP"

Screen.MousePointer = HOURGLASS
If glbtermopen Then
Data1.ConnectionString = glbAdoIHRAUDIT
Else
Data1.ConnectionString = glbAdoIHRDB
End If

'Call DropDown_Controls

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call ST_UPD_MODE(False)

Call INI_Controls(Me)
If glbCompSerial = "S/N - 2382W" Then 'Namasco Ltd
    lblTitle(10).Caption = "Pay Rule"
    lblTitle(11).Caption = "Labour Account Number"
    lblTitle(12).Caption = "Restriction Profile"
    'If Len(txtUValue(2).Text) = 0 Then
    '    txtUValue(2).Text = "End of Day"
    'End If
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
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
End Sub

'Private Sub imgEmail_Click(Index As Integer)
'    Call txtEEMail_DblClick(Index)
'End Sub

Private Sub lblEEID_Change()


If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = "Employee ADP Data - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblEEnum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub



Private Sub medCTele_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCTele_ValidationError(Index As Integer, InvalidText As String, StartPosition As Integer)
    glbBYPASS380 = True
End Sub
Public Sub cmdNew_Click()
Dim X As Integer
fglbNew = True

Call SET_UP_MODE

On Error GoTo AddN_Err

For X = 0 To 4
    oCP(X) = ""
    If X <> 4 Then oDC(X) = ""
    oUFN(X) = ""
    oUFV(X) = ""
Next X
xAction = "A"   '24June99 js - added from VB3
Call Set_Control("B", Me)
rsDATA.AddNew


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Exit Sub
AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRDEPEND", "Add")
Call RollBack

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

'txtCP(0).Enabled = TF
'txtCP(1).Enabled = TF
'txtCP(2).Enabled = TF
'txtCP(3).Enabled = TF
'txtCP(4).Enabled = TF
'txtDCP(0).Enabled = TF
'txtDCP(1).Enabled = TF
'txtDCP(2).Enabled = TF
'txtDCP(3).Enabled = TF

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



                   
Function chkEmpOther()
Dim X%

On Error GoTo chkEmpOther_Err

chkEmpOther = False


'If Len(dlpPassDate) > 0 Then
'    If Not IsDate(dlpPassDate) Then
'        MsgBox "Invalid Password Expiration Date"
'        dlpPassDate.SetFocus
'        Exit Function
'    End If
'End If


chkEmpOther = True

Exit Function

chkEmpOther_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEmpOther", "HREMP", "edit/Add")
Call RollBack '17Aug99 js

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


Private Sub txtDOR2ADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDORADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEDoctor_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEEMail_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub
'Private Function FldList()
'Dim SQLQ
''SQLQ = ""
''SQLQ = SQLQ & "ER_COMPNO,ER_EMPNBR, ER_CITIZENSHIP, ER_PASSCOUNTRY, ER_PASSCOUNTRYDATE, "
''SQLQ = SQLQ & "ER_VISAPERMITNO, ER_VISAPERMITDATE,ER_PASSPORTNO,"
''SQLQ = SQLQ & "ER_LDATE, ER_LTIME, ER_LUSER"
''
''If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
''FldList = SQLQ
'End Function

Public Sub Display_Value()
    Dim SQLQ
    If rsDATA.EOF Or rsDATA.BOF Then
        Call Set_Control("B", Me)
        Call SET_UP_MODE
        Exit Sub
    End If
    
If glbtermopen Then
    SQLQ = "Select * from Term_HR_ADP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select * from HR_ADP "
    SQLQ = SQLQ & " where AP_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
'Data1.RecordSource = SQLQ
'Data1.Refresh
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
Me.cmdModify_Click
If glbCompSerial = "S/N - 2382W" Then 'Namasco Ltd
    If Len(txtUValue(2).Text) = 0 Then
        txtUValue(2).Text = "End of Day"
    End If
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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_ADP_Data
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
Call ST_UPD_MODE(TF)
End Sub

Private Sub txtCP_Change(Index As Integer)
    Call Set_ClockFull_Value
End Sub

Private Sub txtCP_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtDCP_Change(Index As Integer)
    Call Set_DataCtrlFull_Value
End Sub

Private Sub txtDCP_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUName_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDCP_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUValue_Change(Index As Integer)
    Dim i As Integer
    'Vital Aire
    If glbCompSerial = "S/N - 2380W" Then
        If Index = 0 Then
            If cmbStatusFlag1.ListCount > -1 Then
                For i = 1 To cmbStatusFlag1.ListCount - 1
                    If Trim(txtUValue(Index)) = "" Then
                        cmbStatusFlag1.ListIndex = -1
                    Else
                        If Left(cmbStatusFlag1.List(i), InStr(1, cmbStatusFlag1.List(i), " -") - 1) = RTrim(txtUValue(Index)) Then
                            cmbStatusFlag1.ListIndex = i
                            Exit Sub
                        Else
                            cmbStatusFlag1.ListIndex = -1
                        End If
                    End If
                Next i
            End If
        ElseIf Index = 1 Then
            If cmbStatusFlag2.ListCount > -1 Then
                For i = 1 To cmbStatusFlag2.ListCount - 1
                    If Trim(txtUValue(Index)) = "" Then
                        cmbStatusFlag2.ListIndex = -1
                    Else
                        If Left(cmbStatusFlag2.List(i), InStr(1, cmbStatusFlag2.List(i), " -") - 1) = RTrim(txtUValue(Index)) Then
                            cmbStatusFlag2.ListIndex = i
                            Exit Sub
                        Else
                            cmbStatusFlag2.ListIndex = -1
                        End If
                    End If
                Next i
            End If
        ElseIf Index = 2 Then
            If cmbStatusFlag3.ListCount > -1 Then
                For i = 1 To cmbStatusFlag3.ListCount - 1
                    If Trim(txtUValue(Index)) = "" Then
                        cmbStatusFlag3.ListIndex = -1
                    Else
                        If Left(cmbStatusFlag3.List(i), InStr(1, cmbStatusFlag3.List(i), " -") - 1) = RTrim(txtUValue(Index)) Then
                            cmbStatusFlag3.ListIndex = i
                            Exit Sub
                        Else
                            cmbStatusFlag3.ListIndex = -1
                        End If
                    End If
                Next i
            End If
        
        ElseIf Index = 3 Then
            If cmbStatusFlag4.ListCount > -1 Then
                For i = 1 To cmbStatusFlag4.ListCount - 1
                    If Trim(txtUValue(Index)) = "" Then
                        cmbStatusFlag4.ListIndex = -1
                    Else
                        If Left(cmbStatusFlag4.List(i), InStr(1, cmbStatusFlag4.List(i), " -") - 1) = RTrim(txtUValue(Index)) Then
                            cmbStatusFlag4.ListIndex = i
                            Exit Sub
                        Else
                            cmbStatusFlag4.ListIndex = -1
                        End If
                    End If
                Next i
            End If
        ElseIf Index = 4 Then
            If cmbStatusFlag5.ListCount > -1 Then
                For i = 1 To cmbStatusFlag5.ListCount - 1
                    If Trim(txtUValue(Index)) = "" Then
                        cmbStatusFlag5.ListIndex = -1
                    Else
                        If Left(cmbStatusFlag5.List(i), InStr(1, cmbStatusFlag5.List(i), " -") - 1) = RTrim(txtUValue(Index)) Then
                            cmbStatusFlag5.ListIndex = i
                            Exit Sub
                        Else
                            cmbStatusFlag5.ListIndex = -1
                        End If
                    End If
                Next i
            End If
        End If
    End If
End Sub

Private Sub txtUValue_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function AUDITBENF(ACTX, aType)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim xFlag As Boolean
Dim strFields As String
Dim X As Integer

On Error GoTo AUDIT_ERR
AUDITBENF = False

xFlag = False
For X = 0 To 4
    If oCP(X) <> txtCP(X) Then xFlag = True
    If X <> 4 Then
        If oDC(X) <> txtDCP(X) Then xFlag = True
    End If
    'If oUFN(x) <> txtUName(x) Then xFlag = True
    If oUFV(X) <> txtUValue(X) Then xFlag = True
Next X
If xFlag = False Then GoTo MODNOUPD

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

'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function


Private Sub ADP_DropDown_Controls()

Dim xDiv

'Vital Aire
If glbCompSerial = "S/N - 2380W" Then
    xDiv = GetEmpData(glbLEE_ID, "ED_DIV")
    
    cmbClockFull.Visible = True
    cmbDataCtrlFull.Visible = True
    lblTitle(15).Visible = True
    lblTitle(16).Visible = True
    
    txtUValue(0).Visible = False
    txtUValue(1).Visible = False
    txtUValue(2).Visible = False
    txtUValue(3).Visible = False
    txtUValue(4).Visible = False
    
    cmbClockFull.Clear
    cmbDataCtrlFull.Clear
    cmbStatusFlag1.Clear
    cmbStatusFlag2.Clear
    cmbStatusFlag3.Clear
    cmbStatusFlag4.Clear
    cmbStatusFlag5.Clear
    
    'Company PWN - Vital Aire
    'If xDiv = "PWN" Then
    'Ticket #21702 Franks 03/07/2012
    'Ticket #23017 - "D2X"
    If xDiv = "PWN" Or xDiv = "MZT" Or xDiv = "D2X" Then
        cmbStatusFlag1.Visible = True
        cmbStatusFlag2.Visible = True
        cmbStatusFlag3.Visible = True
        cmbStatusFlag4.Visible = True
        cmbStatusFlag5.Visible = False
        
        lblTitle(10).Visible = True
        lblTitle(10).Caption = "Status Flag 1"
        lblTitle(11).Visible = True
        lblTitle(11).Caption = "Status Flag 2"
        lblTitle(12).Visible = True
        lblTitle(12).Caption = "Status Flag 5"
        lblTitle(13).Visible = True
        lblTitle(13).Caption = "Status Flag 6"
        lblTitle(14).Visible = False
    
        cmbClockFull.AddItem ""
        cmbClockFull.AddItem "  02P - Stat Pay"
        cmbClockFull.AddItem "  04P - 4% Vacation Pay"
        cmbClockFull.AddItem "  06P - 6% Holiday Pay"
        cmbClockFull.AddItem "  08P - 8% Holiday Pay"
        cmbClockFull.AddItem "  09P - 9% Holiday Pay"
        cmbClockFull.AddItem "  10P - 10% Holiday Pay"
        cmbClockFull.AddItem "  12P - 12% Holiday Pay"
        cmbClockFull.AddItem " D - BC MED DEPENDENT"
        cmbClockFull.AddItem " F - BC MED FAMILY"
        cmbClockFull.AddItem " S - BC MED SINGLE"
        cmbClockFull.AddItem "6 - AHC 60% SUBSIDY"
        cmbClockFull.AddItem "F - AHC FAMILY"
        cmbClockFull.AddItem "S - AHC SINGLE"
        
        cmbDataCtrlFull.AddItem ""
        cmbDataCtrlFull.AddItem "F - Full Rate - CASUAL"
        cmbDataCtrlFull.AddItem "R - Reduced - BENEFITS"
        cmbDataCtrlFull.AddItem "F 7 - Full Local 527" 'Ticket #13834
        
        cmbStatusFlag1.AddItem ""
        cmbStatusFlag1.AddItem "P - Pieceworkers"
        cmbStatusFlag1.AddItem "V - BC UNION VACATION AC"
        
        cmbStatusFlag2.AddItem ""
        cmbStatusFlag2.AddItem "2 - 4.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "3 - 69.33HRS 8.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "4 - 8.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "5 - 10.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "6 - 12.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "7 - 8.00% Vacation Accrual"
        cmbStatusFlag2.AddItem "V - Negotiated 4wks Vacation"
        
        cmbStatusFlag3.AddItem ""
        cmbStatusFlag3.AddItem "U - Union dues for 213"
        
        cmbStatusFlag4.AddItem ""
        cmbStatusFlag4.AddItem "S - STD Qualifier"
        
    ElseIf xDiv = "XTT" Then
        cmbStatusFlag1.Visible = True
        cmbStatusFlag2.Visible = True
        cmbStatusFlag3.Visible = True
        cmbStatusFlag4.Visible = False
        cmbStatusFlag5.Visible = False
    
        lblTitle(10).Visible = True
        lblTitle(10).Caption = "Status Flag 1"
        lblTitle(11).Visible = True
        lblTitle(11).Caption = "Status Flag 2"
        lblTitle(12).Visible = True
        lblTitle(12).Caption = "Status Flag 6"
        lblTitle(13).Visible = False
        lblTitle(14).Visible = False
    
        cmbClockFull.AddItem ""
        cmbClockFull.AddItem "A - Apprentice"
        cmbClockFull.AddItem "  04 - 4% Vac"
        cmbClockFull.AddItem "  08 - 8% Vac"
        cmbClockFull.AddItem "F - Foreman"
        cmbClockFull.AddItem "J - Journeyman"
        cmbClockFull.AddItem "J71 - Union 45 per hr"
        
        cmbDataCtrlFull.AddItem ""
        cmbDataCtrlFull.AddItem " 4 - 4% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem " 6 - 6% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem " 8 - 8% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem "F - 1.4"
        cmbDataCtrlFull.AddItem "F 1 - Full-local 800"
        cmbDataCtrlFull.AddItem "F 2 - Full-local 046"
        'Ticket #13409 - Begin Frank Jul 23, 2007
        cmbDataCtrlFull.AddItem "F 3 - Full-local 463"
        cmbDataCtrlFull.AddItem "F 4 - FULL/QC 144,825"
        cmbDataCtrlFull.AddItem "F 5 - Full-local 67/599"
        cmbDataCtrlFull.AddItem "F 6 - Full-local 71 Vac Ac"
        'Ticket #13409 - End
        cmbDataCtrlFull.AddItem "F4 - FULL -4% VAC ACCRL"
        cmbDataCtrlFull.AddItem "F8 - FULL -8% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "R - 1.248"
        cmbDataCtrlFull.AddItem "R4 - REDUCED -4% Vac Accrl"
        cmbDataCtrlFull.AddItem "R6 - VACATION 6% ACCRUAL"
        cmbDataCtrlFull.AddItem "R8 - REDUCED -8% Vac Accrl"
        
        cmbStatusFlag1.AddItem ""
        cmbStatusFlag1.AddItem "6 - Do not use 11% Vac Pay for 7"
        cmbStatusFlag1.AddItem "8 - 11% Vac Payout"
        
        cmbStatusFlag2.AddItem ""
        cmbStatusFlag2.AddItem "Q - 11.5% Vac Pay for QC 71"
        
        cmbStatusFlag3.AddItem ""
        cmbStatusFlag3.AddItem "S - STD EARNINGS"
        
    ElseIf xDiv = "XTM" Then
        cmbStatusFlag1.Visible = True
        cmbStatusFlag2.Visible = False
        cmbStatusFlag3.Visible = False
        cmbStatusFlag4.Visible = False
        cmbStatusFlag5.Visible = False
        
        lblTitle(10).Visible = True
        lblTitle(10).Caption = "Status Flag 1"
        lblTitle(11).Visible = False
        lblTitle(12).Visible = False
        lblTitle(13).Visible = False
        lblTitle(14).Visible = False
        
        cmbClockFull.AddItem ""
        cmbClockFull.AddItem "  04P - VACATION PAYOUT"
        cmbClockFull.AddItem "  08P - VAC PAYOUT"
        cmbClockFull.AddItem "  10P - VAC PAYOUT"
        cmbClockFull.AddItem "  12P - VAC PAYOUT"
        cmbClockFull.AddItem " D - BC MEDICAL DEPENDENT"
        cmbClockFull.AddItem " S - BC MED SINGLE"
        cmbClockFull.AddItem "A - AHC SINGLE"
        cmbClockFull.AddItem "H - AHC FAMILY"
        cmbClockFull.AddItem "J - Journeyman"
        cmbClockFull.AddItem "JS - JOURNEY/BC SINGLE"
        cmbClockFull.AddItem "S - BC MED SINGLE"
        
        cmbDataCtrlFull.AddItem ""
        cmbDataCtrlFull.AddItem " 4 - 4% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem " 6 - 6% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem " 8 - 8% VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem " V - VACATION ACCRUAL"
        cmbDataCtrlFull.AddItem "F - Full Rate"
        cmbDataCtrlFull.AddItem "F0 - FULL-10% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "F4 - FULL -4% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "F8 - FULL -8% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "R - Reduced Rate"
        cmbDataCtrlFull.AddItem "R4 - REDUCED-4% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "R6 - REDUCED-6% VAC ACCRUAL"
        cmbDataCtrlFull.AddItem "R8 - REDUCED -8% VAC ACCRUAL"
        
        cmbStatusFlag1.AddItem ""
        cmbStatusFlag1.AddItem "U - Union EE for dues"
        
    Else
        cmbClockFull.Visible = False
        cmbDataCtrlFull.Visible = False
        lblTitle(15).Visible = False
        lblTitle(16).Visible = False
    
        txtUValue(0).Visible = True
        txtUValue(1).Visible = True
        txtUValue(2).Visible = True
        txtUValue(3).Visible = True
        txtUValue(4).Visible = True
        
        cmbStatusFlag1.Visible = False
        cmbStatusFlag2.Visible = False
        cmbStatusFlag3.Visible = False
        cmbStatusFlag4.Visible = False
        cmbStatusFlag5.Visible = False
    End If
Else
    cmbClockFull.Visible = False
    cmbDataCtrlFull.Visible = False
    lblTitle(15).Visible = False
    lblTitle(16).Visible = False
End If

End Sub

Private Sub Emp_ADP_ClockFull()
    Dim xClockCode As String
    
    If cmbClockFull.ListIndex > -1 Then
        If Len(cmbClockFull.Text) > 0 Then
            xClockCode = Left(cmbClockFull.Text, InStr(1, cmbClockFull.Text, " -") - 1)
            If Len(xClockCode) > 0 Then
                txtCP(0).Text = IIf(Left(xClockCode, 1) = " ", "", Left(xClockCode, 1))
                txtCP(1).Text = IIf(Mid(xClockCode, 2, 1) = " ", "", Mid(xClockCode, 2, 1))
                txtCP(2).Text = IIf(Mid(xClockCode, 3, 1) = " ", "", Mid(xClockCode, 3, 1))
                txtCP(3).Text = IIf(Mid(xClockCode, 4, 1) = " ", "", Mid(xClockCode, 4, 1))
                txtCP(4).Text = IIf(Mid(xClockCode, 5, 1) = " ", "", Mid(xClockCode, 5, 1))
            Else
                txtCP(0).Text = ""
                txtCP(1).Text = ""
                txtCP(2).Text = ""
                txtCP(3).Text = ""
                txtCP(4).Text = ""
            End If
        Else
            txtCP(0).Text = ""
            txtCP(1).Text = ""
            txtCP(2).Text = ""
            txtCP(3).Text = ""
            txtCP(4).Text = ""
        End If
    End If
End Sub

Private Sub Emp_ADP_DataCtrlFull()
    Dim xDataCtrlCode As String
    
    If cmbDataCtrlFull.ListIndex > -1 Then
        If Len(cmbDataCtrlFull.Text) > 0 Then
            xDataCtrlCode = Left(cmbDataCtrlFull.Text, InStr(1, cmbDataCtrlFull.Text, " -") - 1)
            If Len(xDataCtrlCode) > 0 Then
                txtDCP(0).Text = IIf(Left(xDataCtrlCode, 1) = " ", "", Left(xDataCtrlCode, 1))
                txtDCP(1).Text = IIf(Mid(xDataCtrlCode, 2, 1) = " ", "", Mid(xDataCtrlCode, 2, 1))
                txtDCP(2).Text = IIf(Mid(xDataCtrlCode, 3, 1) = " ", "", Mid(xDataCtrlCode, 3, 1))
                txtDCP(3).Text = IIf(Mid(xDataCtrlCode, 4, 1) = " ", "", Mid(xDataCtrlCode, 4, 1))
            Else
                txtDCP(0).Text = ""
                txtDCP(1).Text = ""
                txtDCP(2).Text = ""
                txtDCP(3).Text = ""
            End If
        Else
            txtDCP(0).Text = ""
            txtDCP(1).Text = ""
            txtDCP(2).Text = ""
            txtDCP(3).Text = ""
        End If
    End If
End Sub

Private Sub Set_ClockFull_Value()
    Dim xClockCode As String
    Dim i As Integer
    
    xClockCode = ""
    If txtCP(0).Text <> "" Then
        xClockCode = xClockCode & txtCP(0).Text
    Else
        xClockCode = xClockCode & " "
    End If
    If txtCP(1).Text <> "" Then
        xClockCode = xClockCode & txtCP(1).Text
    Else
        xClockCode = xClockCode & " "
    End If
    If txtCP(2).Text <> "" Then
        xClockCode = xClockCode & txtCP(2).Text
    Else
        xClockCode = xClockCode & " "
    End If
    If txtCP(3).Text <> "" Then
        xClockCode = xClockCode & txtCP(3).Text
    Else
        xClockCode = xClockCode & " "
    End If
    If txtCP(4).Text <> "" Then
        xClockCode = xClockCode & txtCP(4).Text
    Else
        xClockCode = xClockCode & " "
    End If
    If cmbClockFull.ListCount > -1 Then
        For i = 1 To cmbClockFull.ListCount - 1
            If Trim(xClockCode) = "" Then
                cmbClockFull.ListIndex = -1
            Else
                If Left(cmbClockFull.List(i), InStr(1, cmbClockFull.List(i), " -") - 1) = RTrim(xClockCode) Then
                    cmbClockFull.ListIndex = i
                    Exit Sub
                Else
                    cmbClockFull.ListIndex = -1
                End If
            End If
        Next i
    End If
End Sub

Private Sub Set_DataCtrlFull_Value()
    Dim xDataCtrlCode As String
    Dim i As Integer
    
    xDataCtrlCode = ""
    If txtDCP(0).Text <> "" Then
        xDataCtrlCode = xDataCtrlCode & txtDCP(0).Text
    Else
        xDataCtrlCode = xDataCtrlCode & " "
    End If
    If txtDCP(1).Text <> "" Then
        xDataCtrlCode = xDataCtrlCode & txtDCP(1).Text
    Else
        xDataCtrlCode = xDataCtrlCode & " "
    End If
    If txtDCP(2).Text <> "" Then
        xDataCtrlCode = xDataCtrlCode & txtDCP(2).Text
    Else
        xDataCtrlCode = xDataCtrlCode & " "
    End If
    If txtDCP(3).Text <> "" Then
        xDataCtrlCode = xDataCtrlCode & txtDCP(3).Text
    Else
        xDataCtrlCode = xDataCtrlCode & " "
    End If
    If cmbDataCtrlFull.ListCount > -1 Then
        For i = 1 To cmbDataCtrlFull.ListCount - 1
            If Trim(xDataCtrlCode) = "" Then
                cmbDataCtrlFull.ListIndex = -1
            Else
                If Left(cmbDataCtrlFull.List(i), InStr(1, cmbDataCtrlFull.List(i), " -") - 1) = RTrim(xDataCtrlCode) Then
                    cmbDataCtrlFull.ListIndex = i
                    Exit Sub
                Else
                    cmbDataCtrlFull.ListIndex = -1
                End If
            End If
        Next i
    End If
End Sub

