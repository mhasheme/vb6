VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEHSINJURY 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Injury/Location "
   ClientHeight    =   9780
   ClientLeft      =   330
   ClientTop       =   810
   ClientWidth     =   12150
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
   ScaleHeight     =   9780
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsinjr.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "fehsinjr.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   11895
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   12150
      _Version        =   65536
      _ExtentX        =   21431
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
         Left            =   6960
         TabIndex        =   54
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   1320
         TabIndex        =   25
         Top             =   135
         Width           =   1245
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
         Left            =   2880
         TabIndex        =   24
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LDATE"
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
      Left            =   2760
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LTIME"
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
      Left            =   4560
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
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
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel frmDetails 
      Height          =   5895
      Left            =   0
      TabIndex        =   29
      Top             =   2970
      Width           =   11955
      _Version        =   65536
      _ExtentX        =   21087
      _ExtentY        =   10398
      _StockProps     =   15
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.TextBox txtOSHACOM 
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
         Height          =   285
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "00-Form 7 sec 6/ OSHA Comment"
         Top             =   4680
         Visible         =   0   'False
         Width           =   6570
      End
      Begin VB.TextBox txtOSHA300 
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
         Height          =   285
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "00-Form 7/OSHA 300"
         Top             =   5400
         Visible         =   0   'False
         Width           =   6570
      End
      Begin VB.CheckBox chkCompleted 
         Caption         =   "Check1"
         DataField       =   "EC_INJURED_ONLINE"
         Height          =   195
         Left            =   3060
         TabIndex        =   48
         Tag             =   "Completed"
         Top             =   105
         Width           =   285
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_JBCODE"
         Height          =   285
         Index           =   12
         Left            =   7560
         TabIndex        =   18
         Tag             =   "01-Position code"
         Top             =   2640
         Visible         =   0   'False
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECJB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_LOC"
         Height          =   285
         Index           =   8
         Left            =   7575
         TabIndex        =   16
         Tag             =   "00-Location of Incident"
         Top             =   2280
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_SECONDARY"
         Height          =   285
         Index           =   7
         Left            =   7575
         TabIndex        =   14
         Tag             =   "00-Secondary Cause of Injury"
         Top             =   1920
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_SBODY"
         Height          =   285
         Index           =   3
         Left            =   7575
         TabIndex        =   9
         Tag             =   "00-Enter Body Site - Code"
         Top             =   1200
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_SFACT"
         Height          =   285
         Index           =   10
         Left            =   7575
         TabIndex        =   7
         Tag             =   "00-Enter Facet - Code"
         Top             =   840
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECFA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_SCODE"
         Height          =   285
         Index           =   11
         Left            =   7560
         TabIndex        =   5
         Tag             =   "00-Injury - Code"
         Top             =   480
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_EQUIP"
         Height          =   285
         Index           =   6
         Left            =   2145
         TabIndex        =   17
         Tag             =   "00-Equipment being used when injured"
         Top             =   2640
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECEQ"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_AREA"
         Height          =   285
         Index           =   5
         Left            =   2145
         TabIndex        =   15
         Tag             =   "01-Area where incident occurred"
         Top             =   2280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_CAUSECD"
         Height          =   285
         Index           =   4
         Left            =   2145
         TabIndex        =   13
         Tag             =   "01-Primary Cause of Injury"
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_PBODY"
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Tag             =   "01-Enter Body Site - Code"
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_PFACT"
         Height          =   285
         Index           =   9
         Left            =   2145
         TabIndex        =   6
         Tag             =   "01-Enter Facet - Code"
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECFA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_CODE"
         Height          =   285
         Index           =   1
         Left            =   2145
         TabIndex        =   4
         Tag             =   "01-Injury - Code"
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   "Injury"
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
         Left            =   120
         TabIndex        =   1
         Tag             =   "40-Injury"
         Top             =   60
         Width           =   855
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         DataField       =   "EC_COMMENTS"
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
         Height          =   1230
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Tag             =   "00-General Comments"
         Top             =   3360
         Width           =   8205
      End
      Begin VB.TextBox txtTask 
         Appearance      =   0  'Flat
         DataField       =   "EC_TASK"
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
         Left            =   2460
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "01-Task being performed when injured"
         Top             =   1560
         Width           =   6315
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   " Other"
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
         Left            =   2100
         TabIndex        =   3
         Tag             =   "40-Other"
         Top             =   60
         Width           =   1035
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   "Disease"
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
         Left            =   1020
         TabIndex        =   2
         Tag             =   "40-Disease"
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label lblOSHACOM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Form 7 sec 6/ OSHA Comment"
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
         Left            =   210
         TabIndex        =   56
         Top             =   4725
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label lblOSHA300 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Form 7/OSHA 300"
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
         Left            =   210
         TabIndex        =   55
         Top             =   5445
         Visible         =   0   'False
         Width           =   1635
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
         Height          =   255
         Left            =   1200
         TabIndex        =   53
         Top             =   5040
         Width           =   2295
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
         Height          =   255
         Left            =   210
         TabIndex        =   52
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblUpdDateDesc 
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
         Height          =   255
         Left            =   5160
         TabIndex        =   51
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
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
         Left            =   4080
         TabIndex        =   50
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Injured on Line"
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
         Left            =   3360
         TabIndex        =   49
         Top             =   90
         Width           =   1050
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "OH&&S Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5880
         TabIndex        =   47
         Top             =   2685
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblSINJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EC_SINJ"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4620
         TabIndex        =   45
         Top             =   150
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         Left            =   210
         TabIndex        =   44
         Top             =   3110
         Width           =   735
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
         Index           =   1
         Left            =   5880
         TabIndex        =   43
         Top             =   2325
         Width           =   615
      End
      Begin VB.Label lblBody2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Body Site"
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
         Left            =   5880
         TabIndex        =   41
         Top             =   1245
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Facet"
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
         Left            =   5880
         TabIndex        =   40
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Injury"
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
         Left            =   5880
         TabIndex        =   39
         Top             =   525
         Width           =   1185
      End
      Begin VB.Label lblEquipment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment"
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
         Left            =   210
         TabIndex        =   38
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label lblPlantArea 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Area"
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
         Left            =   210
         TabIndex        =   37
         Top             =   2325
         Width           =   900
      End
      Begin VB.Label lblCause 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Cause"
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
         Left            =   210
         TabIndex        =   36
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label lblTask 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Task"
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
         Left            =   210
         TabIndex        =   35
         Top             =   1605
         Width           =   2205
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblBody1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Body Site"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Facet"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Injury"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   32
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label lblIncidentNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "EC_CASE"
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
         Left            =   7920
         TabIndex        =   31
         Top             =   105
         Width           =   90
      End
      Begin VB.Label lblIncident 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5880
         TabIndex        =   30
         Top             =   105
         Width           =   1410
      End
      Begin VB.Label lblSecCause 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Cause"
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
         Left            =   5880
         TabIndex        =   42
         Top             =   1965
         Width           =   1500
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   46
      Top             =   9120
      Width           =   12150
      _Version        =   65536
      _ExtentX        =   21431
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   10680
         Top             =   360
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Caption         =   "Ado2"
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
         Left            =   10500
         Top             =   90
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
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EC_EMPNBR"
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
      Left            =   1650
      TabIndex        =   27
      Top             =   8910
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "001"
      DataField       =   "EC_COMPNO"
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
      Left            =   390
      TabIndex        =   28
      Top             =   8910
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmEHSINJURY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbUpd%
Dim fglbNew
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim glbOccDate
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control



Private Function chkHSInjury()
Dim SQLQ As String, Msg As String, dd#, X%

chkHSInjury = False

On Error GoTo chkHSInjury_Err

If Len(clpCode(1).Text) < 1 Then
    If Not OptInjDis(2).Value Then
        MsgBox "Injury Code is a required field"
        clpCode(1).SetFocus
        Exit Function
    End If
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Injury code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(2).Text) < 1 Then
    If Not OptInjDis(2).Value Then
        MsgBox "#1 Body Site is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
Else
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "#1 Body Site code must be valid"
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(3).Text) > 0 Then
  If clpCode(3).Caption = "Unassigned" Then
    MsgBox "#2 Body Site code must be valid"
    clpCode(3).SetFocus
    Exit Function
  End If
  ' dkostka - 10/02/2001 - Removed on request of Linda
'  If clpCode(3) = clpCode(2) Then
'    MsgBox "#2 Body Site can not be the same as #1 Body Site."
'    clpCode(3).SetFocus
'    Exit Function
'  End If
End If

'Jerry said not to make Task mandatory - like Form 7. Since Linamar has custom code, leaving it
'mandatory for them
'If Not glbWFC Then
If glbLinamar Then
    If Len(txtTask) < 1 And Not OptInjDis(2).Value Then
        MsgBox "Task is a required field"
        txtTask.SetFocus
        Exit Function
    End If
End If

If Len(clpCode(4).Text) < 1 Then
    If Not OptInjDis(2).Value Then
        If glbWFC Then
          MsgBox "Pers Eq #1 is a required field"
            clpCode(4).SetFocus
            Exit Function
        'Else       'As per Next Release Documentation
        '  MsgBox "Primary cause is a required field"
        End If
        'MsgBox "Primary cause is a required field"
    End If
Else
    If clpCode(4).Caption = "Unassigned" Then
        If glbWFC Then
          MsgBox "Pers Eq #1 must be valid"
        Else
          MsgBox "Cause code must be valid"
        End If
        clpCode(4).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(5).Text) < 1 Then
    If Not OptInjDis(2).Value Then
        'MsgBox "Plant Area code is a required field"   'As per Next Release Documentation
        'clpCode(5).SetFocus
        'Exit Function
    End If
    If glbWFC Then
        MsgBox "Plant Area is a required field"
        clpCode(5).SetFocus
        Exit Function
    End If
Else
    If Len(clpCode(5).Text) > 1 Then
        If clpCode(5).Caption = "Unassigned" Then
            MsgBox "Plant Area code must be valid"
            clpCode(5).SetFocus
            Exit Function
        End If
    End If
End If

If Len(clpCode(9).Text) < 1 And Not OptInjDis(2).Value Then
    MsgBox "Facet is a required field"
    clpCode(9).SetFocus
    Exit Function
End If

For X% = 6 To 12
    If clpCode(X%).Visible Then
        If Len(clpCode(X%).Text) > 0 Then
            If clpCode(X%).Caption = "Unassigned" Then
                MsgBox "Invalid code Entered"
                clpCode(X%).SetFocus
                Exit Function
            End If
        End If
    End If
Next

If glbWFC Then
    If Len(clpCode(6)) = 0 Then
        MsgBox "Equipment is a required field"
        clpCode(6).SetFocus
        Exit Function
    End If
    If Len(Trim(clpCode(12).Text)) < 1 Then
        'MsgBox "Job Code is a required field"
        'MsgBox "OH&S Position Code is a required field"
        MsgBox "Position is a required field" 'Ticket #29277 Franks 10/04/2016
        clpCode(12).SetFocus
        Exit Function
    Else
        If clpCode(12).Caption = "Unassigned" Then
            'MsgBox "Job Code must be valid"
            'MsgBox "OH&S Position Code must be valid"
            MsgBox "Position Code must be valid" 'Ticket #29277 Franks 10/04/2016
            clpCode(12).SetFocus
            Exit Function
        End If
    End If
    If Len(txtComments.Text) = 0 Then 'Ticket #29277 Franks 10/04/2016
        MsgBox "Comments is a required field"
        txtComments.SetFocus
        Exit Function
    End If
End If

'If clpCode(1) = clpCode(11) Then
'  MsgBox "#2 Injury Code can not be the same as #1 Injury Code."
'  clpCode(11).SetFocus
'  Exit Function
'End If
'~~~~~~~~~~~~~~commented out by RAUBREY 6/2/97 ~~~~~~~~~~~~~~
'If clpCode(9) = clpCode(10) And Len(clpCode(9)) > 0 Then
'  MsgBox "#2 Facet can not be the same as #1 Facet."
'  clpCode(10).SetFocus
'  Exit Function
'End If

If glbLinamar Then
    'If Len(clpCode(4)) = 0 Then
    '    MsgBox lblCause.Caption & " is a required field"
    '    clpCode(4).SetFocus
    '    Exit Function
    'End If
    If Len(clpCode(5)) = 0 Then
        MsgBox "Plant Area is a required field"
        clpCode(5).SetFocus
        Exit Function
    End If
    If Len(clpCode(6)) = 0 Then
        MsgBox "Equipment is a required field"
        clpCode(6).SetFocus
        Exit Function
    End If
    'Ticket #14666
    'If Len(clpCode(11)) = 0 Then
    '    MsgBox "Secondary Injury is a required field"
    '    clpCode(11).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(10)) = 0 Then
    '    MsgBox "Facet is a required field"
    '    clpCode(10).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(3)) = 0 Then
    '    MsgBox "Body Site is a required field"
    '    clpCode(3).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(8)) = 0 Then
    '    MsgBox lblLocation(1).Caption & " is a required field"
    '    clpCode(8).SetFocus
    '    Exit Function
    'End If
    'If Len(txtComments.Text) = 0 Then  'Ticket #16782
    '    MsgBox "Comments is a required field"
    '    txtComments.SetFocus
    '    Exit Function
    'End If
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If OptInjDis(0) Then lblSINJ = "I"
If OptInjDis(1) Then lblSINJ = "D"
If OptInjDis(2) Then lblSINJ = "O"

chkHSInjury = True
Exit Function

chkHSInjury_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkInjury", "HR_OCC_HEALTH_SAFETY", "edit/Add")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function



'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

'Private Sub cmdCAction_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdCancel_Click()
Dim locDate, X
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
If Not (rsDATA.EOF And rsDATA.BOF) Then rsDATA.CancelUpdate
Call Display_Value
fglbNew = False
'Call ST_UPD_MODE(True)  ' reset screen's attributes
Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OCC_HEALTH_SAFETY", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSINJURY" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

'Private Sub cmdIncident_Click()
'frmEHSINCIDENT.Show
'Unload Me
'End Sub

'Private Sub cmdIncident_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
On Error GoTo Mod_Err
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
'clpCode(1).SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OCC_HEALTH_SAFETY", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
On Error GoTo Add_Err

If Not chkHSInjury() Then Exit Sub
rsDATA.Requery

Call UpdUStats(Me) ' update user's stats (who did it and when)
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
data1.Refresh
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
X = NextFormIF("Injury/Location")
Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OCC_HEALTH_SAFETY", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Injury"
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

RHeading = lblEEName & "'s Injury"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdTCause_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub

'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False
Screen.MousePointer = HOURGLASS
On Error GoTo EERError

If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
data1.RecordSource = SQLQ
data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OCC_HEALTH_SAFETY", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function





End Function

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
' Me.cmdModify_Click
glbOnTop = "FRMEHSINJURY"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSINJURY"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer, X% ' records found
glbOnTop = "FRMEHSINJURY"
fglbUpd% = False

'Data1.DatabaseName = glbIHRDB
If glbtermopen Then
    data1.ConnectionString = glbAdoIHRAUDIT
Else
    data1.ConnectionString = glbAdoIHRDB
End If
Screen.MousePointer = DEFAULT

lblLocation(1).Caption = lStr(lblLocation(1).Caption)

If glbWFC Then
  lblCause = "Pers Eq #1"
  lblSecCause = "Pers Eq #2"
  clpCode(4).TABLTitle = "Personal Equipment 1"
  clpCode(4).Tag = "01-Personal Equipment 1"
  clpCode(7).TABLTitle = "Personal Equipment 2"
  clpCode(7).Tag = "01-Personal Equipment 2"
  vbxTrueGrid.Columns(6).Caption = "Pers Eq #1"
  vbxTrueGrid.Columns(7).Caption = "Pers Eq #2"
  lblPosTitle.Visible = True
  clpCode(12).Visible = True
  lblTask.Font.Bold = False
  lblCause.FontBold = True
  lblPlantArea.FontBold = True
  lblEquipment.FontBold = True
  lblPosTitle.Caption = "Position" 'Ticket #29277 Franks 10/04/2016
  lbltitle(14).FontBold = True 'Ticket #29277 Franks 10/04/2016
End If

If glbCompSerial = "S/N - 2387W" Then  'Bird Packaging Limited 'Ticket #13636
    OptInjDis(0).Caption = "Sudden"
    OptInjDis(2).Caption = "Gradual"
    lbltitle(1).Caption = " Other"
    lblType.Caption = "Accident Type"
    Label1.Caption = "Area of Injury"
    Label2.Caption = "Area of Injury"
    lblBody2.Caption = "N/A"
    lblCause.Caption = "Injury Class #1"
    lblSecCause.Caption = "Injury Class #2"
End If

'Ticket #14573
If glbLinamar Then
    'lblCause.Caption = "Root Cause"
    lblCause.Visible = False
    'clpCode(4).Tag = "01-Root Cause of Injury"
    clpCode(4).Visible = False
    lblSecCause.Visible = False
    clpCode(7).Visible = False
    
    'Ticket #14703
    lblLocation(1).Visible = False
    clpCode(8).Visible = False
    
    'Mandatory fields
    'lblCause.FontBold = True
    lblPlantArea.FontBold = True
    lblEquipment.FontBold = True
    'Label3.FontBold = True     'Ticket #14666
    'Label1.FontBold = True     'Ticket #14666
    'lblBody2.FontBold = True   'Ticket #14666
    'lblLocation(1).FontBold = True 'Ticket #14666
    'lblTitle(14).FontBold = True   'Ticket #16782
    
    'Ticket #15172
    OptInjDis(0).Visible = False
    OptInjDis(1).Visible = False
    OptInjDis(2).Visible = False
    chkCompleted.Visible = False
    lbltitle(1).Visible = False
    'lblOSHA300.Top = lblCause.Top   'Ticket #15172
    'txtOSHA300.Top = clpCode(4).Top 'Ticket #15172
    lblOSHA300.Visible = False  'Ticket #15172
    
    'Hemu
    lblTask.FontBold = True
    lblTask.Caption = "Task (Form 7/OSHA 300)"
    lblOSHACOM.Top = lblCause.Top '2325
    txtOSHACOM.Top = clpCode(4).Top '2280
    lblPlantArea.Top = 2325 '2685
    clpCode(5).Top = 2280   '2640
    
    lblEquipment.Top = 2685 '3030
    clpCode(6).Top = 2640   '3000
    lbltitle(14).Top = 3030 '3430
    txtComments.Top = 3280  '3680
    'Hemu
    
    lblOSHACOM.Visible = True
    txtOSHA300.Visible = False  'Ticket #15172
    txtOSHACOM.Visible = True
    txtOSHA300.DataField = "EC_OSHA300"
    txtOSHACOM.DataField = "EC_OSHACOM"
End If

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False)
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False)
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Injury - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Injury - " & Left$(glbLEE_SName, 5)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call Display_Value
'Call ST_UPD_MODE(False)

If Not gSec_Upd_Health_Safety Then
 '   cmdModify.Enabled = False
End If

Call INI_Controls(Me)
 
If glbLinamar Then
    lblPlantArea.FontBold = True
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
Set frmEHSINJURY = Nothing 'carmen may 00
Call NextForm
End Sub


Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

fglbUpd% = YN

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


fUPMode = TF    ' update mode
frmDetails.Enabled = TF
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
txtComments.Enabled = TF
'vbxTrueGrid.Enabled = FT

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT

'cmdPrint.Enabled = FT

'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdCAction.Enabled = FT
'cmdContact.Enabled = FT
'cmdTCause.Enabled = FT
'cmdWSIB.Enabled = FT

OptInjDis(0).Enabled = TF
OptInjDis(1).Enabled = TF
OptInjDis(2).Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
clpCode(6).Enabled = TF
clpCode(7).Enabled = TF
clpCode(8).Enabled = TF
clpCode(9).Enabled = TF
clpCode(10).Enabled = TF
clpCode(11).Enabled = TF
clpCode(12).Enabled = TF
txtComments.Enabled = TF
txtTask.Enabled = TF
If data1.Recordset.BOF And data1.Recordset.EOF Then 'Add by Frank 8/21/2001
    'cmdModify.Enabled = False
Else
'Me.cmdModify_Click
End If

End Sub


Private Sub lblSINJ_Change()
  If lblSINJ = "I" Then OptInjDis(0) = True
  If lblSINJ = "D" Then OptInjDis(1) = True
  If lblSINJ = "O" Then OptInjDis(2) = True
End Sub

Private Sub OptInjDis_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Text1_Change()

End Sub

'Private Sub clpCode_DblClick(Index As Integer)
'Dim oCode As String, OCodeD As String
'oCode = clpCode(Index)
'OCodeD = clpCode(Index)
'Call Get_Code(CodeCodes(Index, 1), CodeCodes(Index, 2))
'If glbCodeRef Then Call ReCreatSnap(Index)
'If Len(glbCode) < 1 Then
'    clpCode(Index).Text = oCode
'    clpCode(Index).Caption = OCodeD
'Else
'    clpCode(Index).Text = glbCode
'    clpCode(Index).Caption = glbCodeDesc
'    clpCode(Index).Visible = True
'End If
'End Sub
'Private Sub clpCode_GotFocus(Index As Integer)
'clpCode(1).Tag = "01-Injury - Code"
'clpCode(2).Tag = "01-Enter Body Site - Code"
'If glbWFC Then
'  clpCode(4).Tag = "01-Pers Eq #1 of Injury"
'  clpCode(7).Tag = "01-Pers Eq #2 of Injury"
'Else
'  clpCode(4).Tag = "01-Primary Cause of Injury"
'  clpCode(7).Tag = "01-Secondary Cause of Injury"
'End If
'clpCode(5).Tag = "01-Area in plant/building where occurred"
'clpCode(9).Tag = "01-Enter Facet - Code"
'clpCode(11).Tag = "01-Injury - Code"
'Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub clpCode_KeyPress(Index As Integer, KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub


Private Sub txtComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtOSHA300_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOSHACOM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTask_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
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
        
        If glbtermopen Then
            SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
            SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
        Else
            SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
            SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        data1.RecordSource = SQLQ
        data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X%
Call Display_Value


End Sub
Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EC_EMPNBR, EC_CASE, EC_OCCDATE, EC_CODE, "
SQLQ = SQLQ & "EC_PBODY, EC_SBODY, EC_CAUSECD, EC_AREA, EC_EQUIP,"
SQLQ = SQLQ & "EC_TASK, EC_LOC, EC_SECONDARY, EC_COMPNO,"
SQLQ = SQLQ & "EC_COMMENTS, EC_PFACT, EC_SFACT, EC_SCODE, EC_SINJ,"
SQLQ = SQLQ & "EC_INJURED_ONLINE,"
SQLQ = SQLQ & "EC_JBCODE, EC_LDATE , EC_LTIME, EC_LUSER"
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbLinamar Then 'Ticket #15172
    SQLQ = SQLQ & ",EC_OSHA300,EC_OSHACOM"
End If
FldList = SQLQ
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    If data1.Recordset.EOF Or data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_CASE=" & data1.Recordset!EC_CASE
    If glbWFC Then
        'SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
        SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_CASE = " & data1.Recordset!EC_CASE
    If glbWFC Or glbCompSerial = "S/N - 2335W" Then
        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"

    
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
UpdateRight = gSec_Upd_Health_Safety
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
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub
Private Sub lblEEID_Change()

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Injury - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If

Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSINJURY.Caption = "Injury - " & Left$(glbLEE_SName, 5)
        frmEHSINJURY.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
    
End If
End Sub




