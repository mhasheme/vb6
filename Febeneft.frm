VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEBENEFITS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Benefits"
   ClientHeight    =   9720
   ClientLeft      =   300
   ClientTop       =   495
   ClientWidth     =   12615
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9720
   ScaleWidth      =   12615
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel FrmBens 
      Height          =   1395
      Left            =   0
      TabIndex        =   75
      Top             =   2880
      Width           =   9855
      _Version        =   65536
      _ExtentX        =   17383
      _ExtentY        =   2461
      _StockProps     =   15
      ForeColor       =   16711680
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
      Font3D          =   1
      Alignment       =   3
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   1
         Left            =   1995
         TabIndex        =   49
         Tag             =   "40-Beneficiary's Date of Birth"
         Top             =   1470
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpBCODE 
         Height          =   285
         Left            =   1995
         TabIndex        =   45
         Tag             =   "01-Benefit - Code"
         Top             =   375
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNCD"
         MaxLength       =   10
      End
      Begin MSMask.MaskEdBox MedSplitPc 
         Height          =   285
         Left            =   2310
         TabIndex        =   50
         Tag             =   "11-Percentage "
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   "##0.00%"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox comRelation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         TabIndex        =   48
         Tag             =   "Relationship of beneficiary to employee"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtBeneName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2310
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "01-Beneficiary of Benefit"
         Top             =   720
         Width           =   4350
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   3
         Left            =   1995
         TabIndex        =   51
         Tag             =   "40-Beneficiary's Date of Death"
         Top             =   2130
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkSepAgree 
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   450
         _Version        =   65536
         _ExtentX        =   794
         _ExtentY        =   450
         _StockProps     =   78
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
      Begin Threed.SSCheck chkSpouseEnt 
         Height          =   255
         Left            =   5520
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   450
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin INFOHR_Controls.DateLookup dlpSepDate 
         Height          =   285
         Left            =   1995
         TabIndex        =   54
         Tag             =   "40-Beneficiary's Date of Separation "
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   46
         Tag             =   "00-Pension Type Code"
         Top             =   360
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EPTY"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   4
         Left            =   1995
         TabIndex        =   55
         Tag             =   "40-Beneficiary's Date of Birth"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1995
         TabIndex        =   56
         Tag             =   "00-Reason Change"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNRE"
         MaxLength       =   10
      End
      Begin VB.Label lblWFCPenDisp 
         Caption         =   $"Febeneft.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   149
         Top             =   3840
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   52
         Left            =   240
         TabIndex        =   147
         Top             =   3150
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Change"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   53
         Left            =   240
         TabIndex        =   146
         Top             =   3525
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pension Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   54
         Left            =   5760
         TabIndex        =   145
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Separation "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   51
         Left            =   240
         TabIndex        =   144
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Spouse Entitled to pension"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   50
         Left            =   3360
         TabIndex        =   143
         Top             =   2460
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Separation Agreement on file"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   49
         Left            =   240
         TabIndex        =   142
         Top             =   2460
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Death"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   47
         Left            =   240
         TabIndex        =   141
         Top             =   2130
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   20
         Left            =   8040
         TabIndex        =   77
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   76
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label lblRel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "relationship"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4680
         TabIndex        =   70
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   71
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiary's Name"
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
         Index           =   14
         Left            =   240
         TabIndex        =   72
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   73
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit"
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
         Index           =   17
         Left            =   240
         TabIndex        =   74
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame fraCopyBS 
      Caption         =   "Copy Beneficiary"
      Height          =   1695
      Left            =   9600
      TabIndex        =   135
      Top             =   8160
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdBSOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   480
         TabIndex        =   140
         Tag             =   "Save the changes made"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdBSCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   139
         Tag             =   "Cancel the changes made"
         Top             =   1200
         Width           =   735
      End
      Begin INFOHR_Controls.CodeLookup clpBCODT 
         Height          =   285
         Left            =   1755
         TabIndex        =   136
         Tag             =   "01-Benefit - Code"
         Top             =   360
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNCD"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpPenType 
         Height          =   285
         Left            =   1755
         TabIndex        =   137
         Tag             =   "00-Pension Type Code"
         Top             =   720
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EPTY"
         MaxLength       =   10
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Pension Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   55
         Left            =   360
         TabIndex        =   148
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Benefit"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   48
         Left            =   360
         TabIndex        =   138
         Top             =   405
         Width           =   1095
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   67
      Top             =   9270
      Width           =   12615
      _Version        =   65536
      _ExtentX        =   22251
      _ExtentY        =   794
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
      Begin VB.CommandButton cmdBsCoopy 
         Appearance      =   0  'Flat
         Caption         =   "Copy Beneficiary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   134
         Tag             =   "Copy Beneficiary"
         Top             =   0
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton cmdRecalAll 
         Caption         =   "&Recalculate All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3420
         TabIndex        =   59
         Top             =   0
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Data3 
         Height          =   375
         Left            =   4680
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   720
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
      Begin VB.CommandButton cmdRecal 
         Caption         =   "&Recalculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         TabIndex        =   58
         Top             =   0
         Width           =   1275
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Index           =   1
         Left            =   10440
         Top             =   210
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
         GridSource      =   "TblBENS"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.CommandButton cmdBens 
         Appearance      =   0  'Flat
         Caption         =   "&Beneficiary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   57
         Tag             =   "Load Beneficiary screen"
         Top             =   0
         Width           =   1365
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   12615
      _Version        =   65536
      _ExtentX        =   22251
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
         Left            =   9840
         TabIndex        =   133
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Original Hire"
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
         Index           =   35
         Left            =   6180
         TabIndex        =   116
         Top             =   160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblDOH 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DOH"
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
         Left            =   7440
         TabIndex        =   115
         Top             =   137
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Range"
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
         Left            =   8640
         TabIndex        =   114
         Top             =   137
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   2640
         TabIndex        =   64
         Top             =   135
         Width           =   720
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Febeneft.frx":00B8
      Height          =   2085
      Left            =   0
      OleObjectBlob   =   "Febeneft.frx":00CC
      TabIndex        =   0
      Top             =   480
      Width           =   9435
   End
   Begin VB.VScrollBar scrControl 
      Height          =   4845
      LargeChange     =   315
      Left            =   10470
      Max             =   100
      SmallChange     =   315
      TabIndex        =   95
      Top             =   2640
      Width           =   300
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Index           =   0
      Left            =   9960
      Top             =   7770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      GridSource      =   "vbxTrueGrid"
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BF_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   9420
      MaxLength       =   25
      TabIndex        =   60
      Top             =   7530
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BF_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   11280
      MaxLength       =   25
      TabIndex        =   61
      Top             =   6690
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BF_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   12960
      MaxLength       =   25
      TabIndex        =   62
      Top             =   6690
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtDiv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9300
      MaxLength       =   4
      TabIndex        =   94
      Tag             =   "00-Specific Division Desired"
      Top             =   7140
      Visible         =   0   'False
      Width           =   870
   End
   Begin TrueOleDBGrid60.TDBGrid TblBENS 
      Bindings        =   "Febeneft.frx":89E8
      Height          =   2145
      Left            =   0
      OleObjectBlob   =   "Febeneft.frx":89FC
      TabIndex        =   110
      Tag             =   "Listing of Beneficiary"
      Top             =   450
      Width           =   9435
   End
   Begin VB.Frame FrmDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   6705
      Left            =   0
      TabIndex        =   78
      Top             =   4560
      Width           =   10305
      Begin VB.TextBox txtCovType 
         Appearance      =   0  'Flat
         DataField       =   "BF_COVER"
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "00-Type of Coverage (S, F, W, X, etc)"
         Top             =   1290
         Width           =   870
      End
      Begin VB.TextBox txtPayrollID 
         Appearance      =   0  'Flat
         DataField       =   "BF_PAYROLL_ID"
         Height          =   285
         Left            =   7620
         TabIndex        =   36
         Top             =   -120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtPerOrDoll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "BF_PERORDOLL"
         Height          =   285
         Left            =   9000
         TabIndex        =   120
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtDWM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "BF_DWM"
         Height          =   285
         Left            =   9300
         TabIndex        =   119
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ComboBox cmbPerOrDoll 
         Height          =   315
         ItemData        =   "Febeneft.frx":C89C
         Left            =   7620
         List            =   "Febeneft.frx":C8A9
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "40-Select Dallor or Percentage"
         Top             =   930
         Width           =   1215
      End
      Begin VB.TextBox txtSortCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7620
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "00-Sort Code"
         Top             =   1620
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cmbDWM 
         Height          =   315
         ItemData        =   "Febeneft.frx":C8C4
         Left            =   8160
         List            =   "Febeneft.frx":C8D1
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "40-Select Day, Week or Month"
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtRoundFactor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         DataField       =   "BF_ROUND"
         Height          =   225
         Left            =   6060
         TabIndex        =   113
         Top             =   2040
         Visible         =   0   'False
         Width           =   345
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "BF_EDATE"
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   6
         Tag             =   "41-Effective Date of coverage"
         Top             =   930
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1225
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1470
         TabIndex        =   4
         Tag             =   "01-Benefit - Code"
         Top             =   600
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BNCD"
         MaxLength       =   10
      End
      Begin VB.Frame frmReleSalaryINFO 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3915
         Left            =   0
         TabIndex        =   96
         Top             =   2790
         Width           =   9495
         Begin VB.TextBox txtSCert 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5880
            MaxLength       =   25
            TabIndex        =   44
            Tag             =   "00-Spouse's Cert #"
            Top             =   3600
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtSPlan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   43
            Tag             =   "00-Spouse's Plan"
            Top             =   3600
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtSComp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5880
            MaxLength       =   25
            TabIndex        =   42
            Tag             =   "00-Spouse's Company"
            Top             =   3240
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.CheckBox chkSPlan 
            Caption         =   "Spouse's Plan"
            DataField       =   "BF_COORDINATION"
            Height          =   255
            Left            =   0
            TabIndex        =   41
            Tag             =   "40-Spouse's Plan -y/n"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtEmployeeID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7710
            MaxLength       =   12
            TabIndex        =   34
            Tag             =   "00-Employee ID"
            Top             =   1725
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtPreAftTax 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            DataField       =   "BF_PTAX"
            DataSource      =   "Data1"
            Height          =   225
            Left            =   8880
            TabIndex        =   111
            Top             =   360
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.ComboBox comPreAftTax 
            Height          =   315
            ItemData        =   "Febeneft.frx":C8F0
            Left            =   7710
            List            =   "Febeneft.frx":C8FA
            TabIndex        =   24
            Tag             =   "Pre Tax/After Tax"
            Top             =   315
            Width           =   1095
         End
         Begin VB.TextBox txtPer 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BF_PER"
            Enabled         =   0   'False
            Height          =   285
            Left            =   7710
            MaxLength       =   5
            TabIndex        =   21
            Tag             =   "10-Enter the base insurance unit"
            Top             =   0
            Width           =   1050
         End
         Begin VB.TextBox txtTAXBEN 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BF_TAXBEN"
            Height          =   285
            Left            =   7710
            MaxLength       =   1
            TabIndex        =   27
            Tag             =   "00-Taxable Benefit    Y=Yes     N=No"
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtPolicy 
            Appearance      =   0  'Flat
            DataField       =   "BF_POLICY"
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   33
            Tag             =   "00-Policy Number"
            Top             =   1710
            Width           =   4215
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "BF_COMMENTS"
            Height          =   945
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Tag             =   "00-Comments - free form"
            Top             =   2250
            Width           =   8805
         End
         Begin MSMask.MaskEdBox medPPComp 
            DataField       =   "BF_PCC"
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Tag             =   "10-Percentage paid by company"
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "##0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMCCOST 
            DataField       =   "BF_MTHCCOST"
            Height          =   285
            Left            =   1800
            TabIndex        =   25
            Tag             =   "20-Monthly company cost"
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$##,##0.0000;($##,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medUnitCost 
            DataField       =   "BF_UNITCOST"
            Height          =   285
            Left            =   4800
            TabIndex        =   20
            Tag             =   "20-Enter Unit Cost"
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.000000;($#,##0.000000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPPE 
            DataField       =   "BF_PCE"
            Height          =   285
            Left            =   4800
            TabIndex        =   23
            Tag             =   "10-Percentage paid by employee"
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medCompCost 
            DataField       =   "BF_CCOST"
            Height          =   315
            Left            =   1800
            TabIndex        =   28
            Tag             =   "11-Cost of Benefit to Company"
            Top             =   990
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEECost 
            DataField       =   "BF_ECOST"
            Height          =   315
            Left            =   4800
            TabIndex        =   29
            Tag             =   "11-Cost of benefit to Employee"
            Top             =   990
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medTCost 
            DataField       =   "BF_TCOST"
            Height          =   315
            Left            =   7710
            TabIndex        =   30
            Tag             =   "21-Total Cost of the Coverage"
            Top             =   990
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Format          =   "$##,##0.0000;($##,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medCovAmount 
            DataField       =   "BF_AMT"
            Height          =   285
            Left            =   1800
            TabIndex        =   19
            Tag             =   "20-Amount of Coverage"
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medCYTD 
            DataField       =   "BF_CYTD"
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Tag             =   "11-Current YTD Company"
            Top             =   1350
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEYTD 
            DataField       =   "BF_EYTD"
            Height          =   315
            Left            =   4800
            TabIndex        =   32
            Tag             =   "11-Current YTD Employee"
            Top             =   1350
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMECOST 
            DataField       =   "BF_MTHECOST"
            Height          =   285
            Left            =   4800
            TabIndex        =   26
            Tag             =   "20-Monthly employee cost"
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$##,##0.0000;($##,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medRateLevel 
            DataField       =   "BF_RATELEVEL"
            Height          =   285
            Left            =   7710
            TabIndex        =   40
            Tag             =   "10-Rate Level"
            Top             =   1725
            Visible         =   0   'False
            Width           =   480
            _ExtentX        =   847
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
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Taxable Benefit"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   24
            Left            =   6300
            TabIndex        =   130
            Top             =   705
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Level"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   46
            Left            =   6300
            TabIndex        =   129
            Top             =   1725
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse's Cert #"
            Height          =   255
            Index           =   43
            Left            =   4080
            TabIndex        =   127
            Top             =   3600
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse's Plan"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   126
            Top             =   3600
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse's Company"
            Height          =   255
            Index           =   41
            Left            =   4080
            TabIndex        =   125
            Top             =   3240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   39
            Left            =   3360
            TabIndex        =   122
            Top             =   1350
            Width           =   825
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current YTD:  Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   38
            Left            =   0
            TabIndex        =   121
            Top             =   1350
            Width           =   1680
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   34
            Left            =   6300
            TabIndex        =   118
            Top             =   1725
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Tax/After Tax"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   33
            Left            =   6300
            TabIndex        =   112
            Top             =   375
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   22
            Left            =   720
            TabIndex        =   109
            Top             =   660
            Width           =   780
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Paid Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   30
            TabIndex        =   108
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Annual:   Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   107
            Top             =   990
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly: "
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   21
            Left            =   0
            TabIndex        =   106
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Coverage Amount"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   105
            Top             =   30
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cost"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   3360
            TabIndex        =   104
            Top             =   30
            Width           =   795
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Paid Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   3360
            TabIndex        =   103
            Top             =   330
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   23
            Left            =   3360
            TabIndex        =   102
            Top             =   660
            Width           =   825
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   3360
            TabIndex        =   101
            Top             =   990
            Width           =   825
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Per"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   6300
            TabIndex        =   100
            Top             =   45
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   6300
            TabIndex        =   99
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Policy Number"
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   98
            Top             =   1710
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   32
            Left            =   0
            TabIndex        =   97
            Top             =   2010
            Width           =   735
         End
      End
      Begin VB.TextBox txtWaitPeriod 
         Appearance      =   0  'Flat
         DataField       =   "BF_WaitPeriod"
         Height          =   285
         Left            =   7620
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "10-Waiting Period (in months)"
         Top             =   270
         Width           =   480
      End
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   6600
         TabIndex        =   91
         Top             =   1860
         Width           =   2175
         Begin VB.OptionButton optRound 
            Caption         =   "Nearest"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   180
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optRound 
            Caption         =   "Next"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   39
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblRound 
            BackColor       =   &H00E0E0E0&
            DataField       =   "BF_NEXTNEAREST"
            Height          =   465
            Left            =   1800
            TabIndex        =   92
            Top             =   180
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.TextBox txtSalDepn 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         DataField       =   "BF_SALARYDEPENDANT"
         Height          =   225
         Left            =   5820
         TabIndex        =   90
         Top             =   1290
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox comRndFactor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Febeneft.frx":C912
         Left            =   4800
         List            =   "Febeneft.frx":C949
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Tag             =   "Rounding Factor"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.ComboBox comSalDepn 
         Height          =   315
         ItemData        =   "Febeneft.frx":C99D
         Left            =   4800
         List            =   "Febeneft.frx":C9A7
         TabIndex        =   10
         Tag             =   "Salary Dependent"
         Text            =   "No"
         Top             =   1275
         Width           =   855
      End
      Begin MSMask.MaskEdBox medPayPeriodAmount 
         DataField       =   "BF_PPAMT"
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Tag             =   "20-Amount charged for every pay period"
         Top             =   930
         Width           =   975
         _ExtentX        =   1720
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
         Format          =   "#,##0.0000;(#,##0.0000)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox medMaxAmnt 
         DataField       =   "BF_MAXDOL"
         Height          =   285
         Left            =   7620
         TabIndex        =   11
         Tag             =   "20-Enter Maximum Amount"
         Top             =   1290
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox medMinCover 
         DataField       =   "BF_MINIMUM"
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Tag             =   "20-Minimum of Coverage"
         Top             =   1650
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox medMaxCover 
         DataField       =   "BF_MAXIMUM"
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Tag             =   "20-Maximum of Coverage"
         Top             =   1650
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox medSalFactor 
         DataField       =   "BF_FACTOR"
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Tag             =   "10-Salary Factor"
         Top             =   2010
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000000000;(#,##0.000000000)"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Height          =   435
         Left            =   0
         TabIndex        =   88
         Top             =   2280
         Width           =   8775
         Begin VB.OptionButton optActual 
            Caption         =   "   Premium"
            Height          =   255
            Index           =   1
            Left            =   4830
            TabIndex        =   18
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton optActual 
            Caption         =   "   Actual"
            Height          =   255
            Index           =   0
            Left            =   1710
            TabIndex        =   17
            Top             =   150
            Width           =   915
         End
         Begin VB.Label lblAP 
            DataField       =   "BF_PREMIUM"
            Height          =   195
            Left            =   6900
            TabIndex        =   89
            Top             =   150
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin INFOHR_Controls.CodeLookup clpGroup 
         DataField       =   "BF_GROUP"
         Height          =   285
         Left            =   1470
         TabIndex        =   1
         Top             =   270
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   2
         Left            =   7305
         TabIndex        =   5
         Tag             =   "41-Benefit End Date"
         Top             =   600
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1300
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   44
         Left            =   6210
         TabIndex        =   132
         Top             =   600
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Coverage Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   131
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   45
         Left            =   5640
         TabIndex        =   128
         Top             =   -60
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dollar/Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   37
         Left            =   6210
         TabIndex        =   124
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   0
         TabIndex        =   123
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sort code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   36
         Left            =   6240
         TabIndex        =   117
         Top             =   1665
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   6210
         TabIndex        =   93
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rounding Factor"
         Height          =   315
         Index           =   29
         Left            =   3300
         TabIndex        =   87
         Top             =   2010
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Factor"
         Height          =   315
         Index           =   27
         Left            =   0
         TabIndex        =   86
         Top             =   2010
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Coverage"
         Height          =   315
         Index           =   28
         Left            =   3300
         TabIndex        =   85
         Top             =   1650
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Coverage"
         Height          =   315
         Index           =   26
         Left            =   0
         TabIndex        =   84
         Top             =   1650
         Width           =   1515
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Dependent"
         Height          =   315
         Index           =   25
         Left            =   3300
         TabIndex        =   83
         Top             =   1275
         Width           =   1335
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   19
         Left            =   9240
         TabIndex        =   82
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Amount"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   6240
         TabIndex        =   37
         Top             =   1290
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period Amount"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3300
         TabIndex        =   81
         Top             =   930
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit"
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
         Index           =   1
         Left            =   0
         TabIndex        =   80
         Top             =   570
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
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
         Index           =   2
         Left            =   0
         TabIndex        =   79
         Top             =   930
         Width           =   1245
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "BF_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1830
      TabIndex        =   68
      Top             =   6960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "BF_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   69
      Top             =   6960
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEBENEFITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbEmptyNew3 As Integer
Dim OBCode, OCOVER, OTCOST, OPremium, OEDate, OPPE, OPCC, OBenEndDate
Dim OPPAMT, OMAXDOL, OBNAME, OBRELATE, ODOB, OPER, OTOTAL, OMedSplitPc, Actn
Dim OUNITCOST, OBAMT
Dim OMTHCOMP, OMTHEMP, OTAXBEN 'ADDED BY RAUBREY 7/9/97
Dim xED_PT, xED_ORG, xJB_GRPVD, xCovAmt, IfElginLife As Boolean
Dim Flag1, Flag2, Flag3
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsDATA3 As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim FRS1 As ADODB.Recordset
Dim FRS3 As ADODB.Recordset
Dim fglbTB_WP, fglbDWM
Dim fglbNew As Integer
Dim VReturn%
Dim fUpdable
Dim OldGroup, oldCode, OldCovType
Dim fEBGroup
Dim AnCoverAmt
Dim MailBody
Dim locManulifeCertNo As String
Dim locSepFlag As Boolean
Dim NewBNAME
Dim ODateEnd 'Ticket #24275 Franks 08/27/2013
Dim ODateDeath
Dim ODateSepa
Dim xCurrentBens As Boolean
Dim xAsOfBens
Dim xBensChg As Boolean
Dim xPenTypeList 'Ticket #24275 Franks 08/27/2013

Private Sub UpdBeneEndDate(xBEDate, xType, xBCode)
Dim rsTBene As New ADODB.Recordset
Dim SQLQ As String
    If xType = "UPD_END_DATE" Then
        SQLQ = "SELECT * FROM HRBENFT "
        SQLQ = SQLQ & " WHERE BF_EMPNBR = " & glbLEE_ID & " AND (BF_CEASEDATE IS NULL) "
        SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
        rsTBene.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsTBene.EOF
            rsTBene("BF_CEASEDATE") = xBEDate
            rsTBene.Update
            Call AUDIT_MANULIFE_BENF(rsTBene("BF_BCODE"), rsTBene("BF_EDATE"), rsTBene("BF_COVER"), rsTBene("BF_POLICY"), xBEDate)
            rsTBene.MoveNext
        Loop
        rsTBene.Close
    End If
    If xType = "DEL_END_DATE" Then
        If CVDate(Date) < CVDate(xBEDate) Then
            SQLQ = "DELETE FROM HR_MANULIFE_TRAN_AUDIT WHERE MT_EMPNBR = " & glbLEE_ID & " "
            SQLQ = SQLQ & "AND MT_TYPE = 'T' "
            SQLQ = SQLQ & "AND MT_BENEFIT = '" & xBCode & "' "
            SQLQ = SQLQ & "AND MT_CEASEDATE = " & Date_SQL(xBEDate) & " "
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
End Sub

Private Function AUDIT_MANULIFE_BENF(xBCode, xBEDate, xBCover, xPolicy, xBenEndDate) 'No AU_CEASEDATE in HRAUDIT, Jerry said we will add it in next release
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDIT_MANULIFE_BENF = False

If Len(xPolicy) = "" Then
    Exit Function
End If

'BENEFIT End Date
If IsDate(xBenEndDate) Then
    If OBenEndDate = "" Then
        GoTo MODUPD
    Else
        If IsDate(OBenEndDate) Then
            If CVDate(xBenEndDate) <> CVDate(OBenEndDate) Then 'Ticket #15591
                GoTo MODUPD
            End If
        End If
    End If
Else
    If IsDate(OBenEndDate) Then
        GoTo MODUPD
    End If
End If

GoTo MODNOUPD

MODUPD:

rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    rsTB.Close:    GoTo MODNOUPD
End If
If IsNull(rsTB("ED_USER_TEXT1")) Then 'Certificate #
    rsTB.Close:    GoTo MODNOUPD
Else
    If Len(Trim(rsTB("ED_USER_TEXT1"))) = 0 Then
        rsTB.Close:    GoTo MODNOUPD
    End If
End If

rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

rsTA.AddNew
rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
rsTA("MT_PT_TABL") = "EDPT"
rsTA("MT_TYPE") = "T"
rsTA("MT_BENEFIT") = xBCode
rsTA("MT_EDATE") = xBEDate
If IsDate(xBenEndDate) Then
rsTA("MT_CEASEDATE") = xBenEndDate
End If
If Len(xBCover) > 0 Then rsTA("MT_COVER") = xBCover
If Len(Trim(xPolicy)) > 0 Then
    rsTA("MT_POLICY_NO") = Trim(xPolicy)
End If
rsTA("MT_COMPNO") = "001"
rsTA("MT_EMPNBR") = glbLEE_ID
rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
rsTA("MT_UPLOAD") = "N"
rsTA("MT_LUSER") = glbUserID
If Not IsDate(xBenEndDate) Then
    rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
Else
    If CVDate(xBenEndDate) < CVDate(Date) Then 'WFC Ticket #14867
        rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
    Else
        rsTA("MT_LDATE") = Format(xBenEndDate, "SHORT DATE")
    End If
End If
rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
rsTA("MT_LTIME") = Time$

rsTA.Update

MODNOUPD:
AUDIT_MANULIFE_BENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT RECORD", "MANULIFE AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function AUDITBENF(ACTX, aType)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR
AUDITBENF = False

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
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST,AU_CEASEDATE, "
strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Then GoTo MODUPD
If aType = 1 Then 'BENEFITS
  If OBCode <> clpCode(1).Text Or OCOVER <> txtCovType Then GoTo MODUPD
  If OTCOST <> medTCost Or OPremium <> lblAP Then GoTo MODUPD
  If OPPE <> medPPE Or OPCC <> medPPComp Then GoTo MODUPD
  If OPPAMT <> medPayPeriodAmount Or OMAXDOL <> medMaxAmnt Then GoTo MODUPD
  If OMTHCOMP <> medMCCOST Or OMTHEMP <> medMECOST Then GoTo MODUPD 'ADDED BY RAUBREY 7/9/97
  If OTAXBEN <> txtTAXBEN Then GoTo MODUPD 'ADDED BY RAUBREY 7/9/97

    ' DK - 03/16/2000 - Removed encryption code
    ' -----
    If OBAMT <> medCovAmount Or OUNITCOST <> medUnitCost Then GoTo MODUPD
    ' -----
  If OPER <> txtPer Or OEDate <> dlpDate(0).Text Then GoTo MODUPD
  If OBenEndDate <> dlpDate(2).Text Then GoTo MODUPD 'Ticket #23594 Franks 04/17/2013
Else
  If OBNAME <> txtBeneName Or OBRELATE <> lblRel Then GoTo MODUPD
  If OBCode <> clpBCODE.Text Or ODOB <> dlpDate(1).Text Then GoTo MODUPD
End If
GoTo MODNOUPD

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then
  If aType = 1 Then
    rsTA("AU_BCODE") = clpCode(1).Text
    If txtCovType <> "" Then rsTA("AU_COVER") = txtCovType
    rsTA("AU_EDATE") = dlpDate(0).Text
    rsTA("AU_MAXDOL") = Val(medMaxAmnt)
    'Frank 01/29/04, ticket #5521
    If IsNumeric(medPayPeriodAmount) Then
        rsTA("AU_PPAMT") = medPayPeriodAmount
    End If
    If IsNumeric(medMCCOST) Then
        rsTA("AU_MTHCCOST") = medMCCOST
    End If
    If IsNumeric(medMECOST) Then
        rsTA("AU_MTHECOST") = medMECOST
    End If
  Else
    rsTA("AU_BCODE") = clpBCODE.Text
    rsTA("AU_BNAME") = txtBeneName
    rsTA("AU_BRELATE") = lblRel
    If IsDate(dlpDate(1).Text) Then         '12Aug99 js
        rsTA("AU_BDOB") = dlpDate(1).Text   '
    Else                              '
        rsTA("AU_BDOB") = Null        '
    End If                            '
    'rsTA("AU_BDOB") = dlpdate(1)     '
    
  End If
Else

  If aType = 1 Then
    If OMTHCOMP <> medMCCOST Then rsTA("AU_MTHCCOST") = medMCCOST 'ADDED BY RAUBREY 7/9/97
    If OMTHEMP <> medMECOST Then rsTA("AU_MTHECOST") = medMECOST 'ADDED BY RAUBREY 7/9/97
    If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN 'ADDED BY RAUBREY 7/9/97
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'changed by raubrey 7/9/97 to make sure benefit code is written
    rsTA("AU_BCODE") = clpCode(1).Text
    If glbWFC Then 'Ticket #13772 Save the old Bcode here since wfc needs the old Bcode.
        If Len(OBCode) > 0 Then
            rsTA("AU_OLDLOC") = Left(OBCode, 10)
            If IsNumeric(OTOTAL) Then
                rsTA("AU_OLDWHRS") = OTOTAL
            End If
        End If
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If OCOVER <> txtCovType Then rsTA("AU_COVER") = txtCovType
    If OTCOST <> medTCost Then rsTA("AU_TCOST") = medTCost
    If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
    If OPPE <> medPPE Then rsTA("AU_PCE") = medPPE
    If OPCC <> medPPComp Then rsTA("AU_PCC") = medPPComp
    If OPPAMT <> medPayPeriodAmount Then
        rsTA("AU_PPAMT") = medPayPeriodAmount
        rsTA("AU_OLDPPMT") = Val(OPPAMT)
    End If
    If OMAXDOL <> medMaxAmnt Then rsTA("AU_MAXDOL") = medMaxAmnt
    If OEDate <> dlpDate(0).Text Then rsTA("AU_EDATE") = dlpDate(0).Text
    If OBenEndDate <> dlpDate(2).Text Then rsTA("AU_CEASEDATE") = IIf(IsDate(dlpDate(2).Text), dlpDate(2).Text, Null)
    If OPER <> txtPer Then rsTA("AU_PER") = txtPer
    ' DK - 03/16/2000 - Removed encryption code
    ' -----
    If OBAMT <> medCovAmount Then rsTA("AU_BAMT") = medCovAmount
    ' -----
    If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)
  Else
    'If OBCODE <>  clpBCODE.text Then rsTA("AU_BCODE") = txtBCODE     'Jaddy 8/6/99
    rsTA("AU_BCODE") = clpBCODE.Text                                  'Jaddy 8/6/99
    If OBNAME <> txtBeneName Then rsTA("AU_BNAME") = txtBeneName
    If OBRELATE <> lblRel Then rsTA("AU_BRELATE") = lblRel
    'If ODOB <> dlpdate(1) Then rsTA("AU_BDOB") = dlpdate(1)
    If IsDate(dlpDate(1).Text) Then             '11Aug99 js
        If ODOB <> dlpDate(1).Text Then         '
            rsTA("AU_BDOB") = dlpDate(1).Text   '
        End If                            '
    Else                                  '
        rsTA("AU_BDOB") = Null            '
    End If                                '
  End If
End If

'If glbSoroc Or glbSyndesis Then
    Dim rsEMP As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
    End If
    rsEMP.Close
'End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
'fRANK 12/15/03,Surrey Place, alway pass Effective Date to LDATE of HRAUDIT
'Frank 01/29/04, as Jerry request, only for new record alway pass Effective Date to LDATE of HRAUDIT
'If glbCompSerial = "S/N - 2347W" Then
'    If FrmDetails.Visible Then
'      rsTA("AU_LDATE") = Format(dlpDate(0).Text, "SHORT DATE")
'    Else
'      rsTA("AU_LDATE") = Date
'    End If
'Else
    If Actn = "A" And FrmDetails.Visible Then
        'If glbWFC Then 'Ticket #14901
        '02/25/10, make this for all customers as Jerry's request
            If CVDate(dlpDate(0).Text) > CVDate(Date) Then
                rsTA("AU_LDATE") = Format(dlpDate(0).Text, "SHORT DATE")
            Else
                rsTA("AU_LDATE") = Date
            End If
        'Else
        '    rsTA("AU_LDATE") = Format(dlpDate(0).Text, "SHORT DATE")
        'End If
    Else
        ''The Walter Fedy Partnership - Ticket #15298
        'If glbCompSerial = "S/N - 2386W" And FrmDetails.Visible Then   'The Walter Fedy Partnership
        If FrmDetails.Visible Then    '02/25/10, make this for all customers as Jerry's request
            If CVDate(dlpDate(0).Text) > CVDate(Date) And Not IsDate(dlpDate(2).Text) Then
                rsTA("AU_LDATE") = Format(dlpDate(0).Text, "SHORT DATE")
            Else
                If IsDate(dlpDate(2).Text) Then
                    If CVDate(dlpDate(2).Text) > CVDate(Date) Then 'Ticket #14867
                        rsTA("AU_LDATE") = Format(dlpDate(2).Text, "SHORT DATE")
                    Else
                        rsTA("AU_LDATE") = Date
                    End If
                ElseIf CVDate(dlpDate(0).Text) > CVDate(Date) Then
                    rsTA("AU_LDATE") = Format(dlpDate(0).Text, "SHORT DATE")
                Else
                    rsTA("AU_LDATE") = Date
                End If
            End If
        Else
            'Ticket #22009 Franks 05/10/2012
            'get the Benefit Start Date, if it future date then use it as LDATE
            'rsTA("AU_LDATE") = Date
            rsTA("AU_LDATE") = getBEDateFrom_clpBCODE(glbLEE_ID, clpBCODE.Text)
        End If
    End If


'End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

If glbWFC Then
    If glbEmpCountry = "CANADA" Then 'Ticket #15818, do not pass benefit to payroll
        Call WFCCNDBeneAuditFlag(glbLEE_ID)
    End If
    If glbEmpCountry = "U.S.A." And clpCode(1).Text = "CC" Then 'Ticket #25307 Franks 04/09/2014
        Call WFCCCBenEndToGTLD
    End If
End If

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function chkEBENEFITS()
Dim SQLQ As String, Msg As String
Dim xlocDays As Integer
Dim xMsg As String

chkEBENEFITS = False
On Error GoTo chkEBENEFITS_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Benefit code is a required field", 64
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Benefit code must be valid", 48
    clpCode(1).SetFocus
    Exit Function
End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Effective Date is not a valid date.", 48
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Effective Date is required field.", 64
    dlpDate(0).SetFocus
    Exit Function
End If
If Len(txtWaitPeriod) > 0 Then
    If IsNumeric(txtWaitPeriod) Then
'        If cmbDWM.ListIndex = -1 Then
'            MsgBox "Please select Day/Week/Month"
'            If cmbDWM.Enabled Then cmbDWM.SetFocus
'            Exit Function
'        End If
    Else
        MsgBox "Waiting Period is invalid", 48
        If txtWaitPeriod.Enabled Then txtWaitPeriod.SetFocus
        Exit Function
    End If
End If
If Len(medCovAmount) > 0 Then
  If Not IsNumeric(medCovAmount) Then
      MsgBox "You Must Enter Dollar Value", 48
      If medCovAmount.Enabled Then medCovAmount.SetFocus
      Exit Function
  End If
End If
' -----

'Comment by Franks Dec 6,02 #3298 , allow the % Company and % Employee to exceed 100%
'If Len(medPPComp) > 0 Then
'    'If Not IsNumeric(medPPComp) Then
'    If Val(medPPComp) < -1 Or Val(medPPComp) > 1 Then
'        MsgBox "You Must Enter Value Between -100% and 100%", 48
'        medPPComp.SetFocus
'        Exit Function
'    End If
'End If

If Len(medPPE) > 0 Then
    If Not IsNumeric(medPPE) Then
        MsgBox "You Must Enter % Value", 48
        medPPE.SetFocus
        Exit Function
    End If
End If

If Len(medPayPeriodAmount) > 0 Then
    If Not IsNumeric(medPayPeriodAmount) Then
        MsgBox "You must enter dollar value"
        medPayPeriodAmount.SetFocus
        Exit Function
    End If
Else    'Hemu 06/18/2003 Begin - if PayPeriodAmount is blank, it gives error in procedure AUDITBENF.
        '                        Only happens if the PayPeriodAmount was previously entered and then editted
        '                        to remove it.
    If medPayPeriodAmount = "" Then
        medPayPeriodAmount = 0
    End If
End If  'Hemu 06/18/2003 End

'Hemu 06/20/2003 Begin - if MaxAmount is blank, it gives error in procedure AUDITBENF.
'                        Only happens if the Max Amount was previously entered and then editted
'                        to remove it.
If medMaxAmnt = "" Then
    medMaxAmnt = 0
End If  'Hemu 06/20/2003 End

'--------added by Jaddy 11/2/99 begin
If comSalDepn = "Yes" Then
    If Len(medMinCover) > 0 Then
        If Not IsNumeric(medMinCover) Then
            MsgBox "Minimum Coverage Must Entry a Number ", 16
            If medMinCover.Enabled Then medMinCover.SetFocus
            Exit Function
        End If
    Else
        medMinCover = 0
    End If
    If Len(medMaxCover) > 0 Then
        If Not IsNumeric(medMaxCover) Then
            MsgBox "Maximum Coverage Must Entry a Number ", 16
            If medMaxCover.Enabled Then medMaxCover.SetFocus
            Exit Function
        Else
            If Val(medMaxCover) > 0 And Val(medMinCover) > 0 Then
                If Val(medMaxCover) < Val(medMinCover) Then
                    MsgBox "Maximum Coverage Must Be Greater Then Minimum Coverage", 16
                    If medMaxCover.Enabled Then medMaxCover.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        medMaxCover = 0
    End If
    If Len(medSalFactor) > 0 Then
        If Not IsNumeric(medSalFactor) Then
            MsgBox "Salary Factor Must Entry a Number ", 16
            If medSalFactor.Enabled Then medSalFactor.SetFocus
            Exit Function
        Else
            If Val(medSalFactor) = 0 Then
                MsgBox "Salary Factor Must Be Greater Then 0", 16
                If medSalFactor.Enabled Then medSalFactor.SetFocus
                Exit Function
            End If
        End If
    Else
          MsgBox "Salary Factor Must Be Greater Then 0", 16
          If medSalFactor.Enabled Then medSalFactor.SetFocus
          Exit Function
    End If

End If
'--------added by Jaddy 11/2/99 end
If optActual(1).Value = True Then
  '~~~~~~~~~~~~~~~~'added by RAUBREY 6/10/97
    If comSalDepn <> "Yes" Then 'Jaddy 11/2/99
        ' DK - 03/16/2000 - Removed encryption code
        ' -----
        If Len(medCovAmount) <= 0 Then
            MsgBox "If Total Cost is Based on Premium Coverage Amount Must Be Entered"
            If medCovAmount.Enabled Then medCovAmount.SetFocus
            Exit Function
        End If
        If Val(medCovAmount) = 0 Then
            MsgBox "Coverage Amount Must Be Greater Then 0", 16
            If medCovAmount.Enabled Then medCovAmount.SetFocus
            Exit Function
        End If
        ' -----
    End If 'Jaddy 11/2/99
    
     If Not IsNumeric(txtPer) Then   'Hemu 05/08/03 Begin
        MsgBox "Per must be numeric"
        txtPer.SetFocus
        Exit Function
    End If  'Hemu 05/08/03 End
    
    If txtPer = 0 Then
    
        MsgBox "If total cost is based on premium, Per must be entered.", 16
        txtPer.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(medUnitCost) Then   'Hemu 05/08/03 Begin
        MsgBox "Unit Cost must be numeric"
        medUnitCost.SetFocus
        Exit Function
    End If  'Hemu 05/08/03 End
    
    If Len(medUnitCost) <= 0 Then
        MsgBox "If total cost is based on premium, Unit Cost must be entered.", 64
        medUnitCost.SetFocus
        Exit Function
    End If
    If medUnitCost = 0 Then
        MsgBox "Unit Cost must be greater then zero.", 16
        medUnitCost.SetFocus
        Exit Function
    End If
    If Len(txtPer) <= 0 Then
        MsgBox "If total cost is based on premium, Per must be greater than zero.", 64
        txtPer.SetFocus
        Exit Function
    End If
End If
If Len(Trim(medPPComp)) = 0 And medPPComp.DataChanged Then medPPComp = 0  'Jaddy 11/3/99

If Len(medPPComp) <= 0 Then
    MsgBox "% Paid By Company Is Required"
    medPPComp.SetFocus
    Exit Function
End If

If Len(medPPE) <= 0 Then
    MsgBox "% Paid By Employee Is Required"
    medPPE.SetFocus
    Exit Function
End If
medTCost = Val(medTCost)


If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Not a valid date."
        dlpDate(1).SetFocus
        Exit Function
    End If
End If


If glbWFC And Len(locManulifeCertNo) > 0 Then 'Ticket #13836
    'Ticket #22038 Franks 05/16/2012
    If fglbNew And Not glbtermopen Then
        If clpCode(1).Text = "HCSA" Or clpCode(1).Text = "HCSA1" Then 'Or clpCode(1).Text = "HCSA4" Or clpCode(1).Text = "HCSA8" Then
            'check if one of this exist
            If isHCSAExist(glbLEE_ID) Then
                MsgBox "Cannot create another HCSA or HCSA1 record if one of them exists." & Chr(10) & "User can change the code between HCSA and HCSA1."
                clpCode(1).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'If clpCode(1).Text = "DENT" Or clpCode(1).Text = "EHC" Or clpCode(1).Text = "HCSA" Or clpCode(1).Text = "HCSA1" Then
    'Ticket #22213 Frank 07/10/2012 add HCSA4 and HCSA8
    If clpCode(1).Text = "DENT" Or clpCode(1).Text = "EHC" Or clpCode(1).Text = "HCSA" Or clpCode(1).Text = "HCSA1" Or clpCode(1).Text = "HCSA4" Or clpCode(1).Text = "HCSA8" Then
        If Len(txtCovType.Text) = 0 Then
            MsgBox "Coverage Type is required for Benefit Code " & clpCode(1).Text
            txtCovType.SetFocus
            Exit Function
        End If
        'Ticket #16286
        If Len(txtPolicy.Text) = 0 Then
            MsgBox "Policy Number is required for Benefit Code " & clpCode(1).Text
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
    
    'Ticket #22009 Franks 05/08/2012
    'Benefit Master - For HCSA & HCSA1 don't allow them to change the Coverage.
    DoEvents
    If clpCode(1).Text = "HCSA" Or clpCode(1).Text = "HCSA1" Then
        If OBCode = clpCode(1).Text Then 'benefit Code not change
            If Len(OCOVER) > 0 Then
                If Not OCOVER = txtCovType.Text Then
                        MsgBox "Can not change Coverage Type for Benefit Code 'HCSA' and 'HCSA1' "
                        txtCovType.SetFocus
                        Exit Function
                End If
            End If
        End If
    End If
    
    'Ticket #13957
    If clpCode(1).Text = "OPLF" Or clpCode(1).Text = "OPLS" Or clpCode(1).Text = "OPLC" Then
        If Len(txtPolicy.Text) = 0 Then
            MsgBox "Policy Number is required for Benefit Code " & clpCode(1).Text
            txtPolicy.SetFocus
            Exit Function
        Else
            If Len(txtPolicy.Text) <> 9 Then
                MsgBox "Invald Policy Number, the format is '#####-###' for Benefit Code " & clpCode(1).Text
                txtPolicy.SetFocus
                Exit Function
            End If
            If Mid(txtPolicy.Text, 6, 1) <> "-" Then
                MsgBox "Invald Policy Number, the format is '#####-###' for Benefit Code " & clpCode(1).Text
                txtPolicy.SetFocus
                Exit Function
            End If
        End If
    End If
End If

'Ticket #22464 - Goodmans
If glbCompSerial = "S/N - 2290W" And Len(medRateLevel.Text) > 0 Then
    If Not IsNumeric(medRateLevel.Text) Then
        MsgBox "Invalid " & lblTitle(46).Caption
        medRateLevel.SetFocus
        Exit Function
    End If
End If

'Frank 11/03/2003 check duplicated benefit code - Begin
If Not glbtermopen Then
Dim rsTBen As New ADODB.Recordset
Dim xFlagBen As Boolean, a%
    SQLQ = "SELECT * FROM HRBENFT "
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & "AND BF_BCODE= '" & clpCode(1) & "' "
    If fglbNew <> True Then
    SQLQ = SQLQ & " AND BF_BENE_ID <> " & rsDATA!BF_BENE_ID
    End If
    rsTBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xFlagBen = False
    If Not rsTBen.EOF Then
        xFlagBen = True
    End If
    rsTBen.Close
    If xFlagBen Then
        Msg = "Duplicate Benefit Code entered. Continue? Yes/No "
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then Exit Function
    End If
End If
'Frank 11/03/2003 check duplicated benefit code - End

'Ticket #22009 Franks 05/11/2012
'DN & EHC's Effective Date can't be less than 30 day's from today with a password being entered
If glbWFC And glbEmpCountry = "CANADA" Then
    'If fglbNew Then
        'If clpCode(1).Text = "DN" Or clpCode(1).Text = "EHC" Then
        'Ticket #26800 Franks 03/13/2015
        If clpCode(1).Text = "DEN" Or clpCode(1).Text = "EHC" Then
            If Len(txtCovType.Text) = 0 Then
                MsgBox "Coverage Type is required for 'DEN' and 'EHC' "
                Exit Function
            Else
                If Not (txtCovType.Text = "S" Or txtCovType.Text = "F" Or txtCovType.Text = "W") Then
                    MsgBox "Invalid Coverage Type. The valid values are 'S', 'F' and 'W' "
                    Exit Function
                End If
            End If
            'Ticket #26800 Franks 03/13/2015 - end
            
            If IsDate(dlpDate(0).Text) Then
                xlocDays = 0
                If fglbNew Then
                    xlocDays = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(Date))
                Else
                    If IsDate(OEDate) Then
                        If Not CVDate(OEDate) = CVDate(dlpDate(0).Text) Then 'Change only
                            xlocDays = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(Date))
                        End If
                    End If
                End If
                'If xlocDays = 0 And xlocDays < 30 Then
                'If xlocDays > 30 Then
                If xlocDays > 90 Then 'Ticket #25116 Franks change 30 days to 90days
                    'xMsg = "Effective Date cannot be less than 30 from TODAY without a password being entered."
                    'Ticket #22108 Franks 07/12/2012
                    xMsg = "Effective Date cannot be more than 90 days prior from today without a password being entered."
                    xMsg = xMsg & Chr(10) & "Please enter a password after click OK button "
                    xMsg = xMsg & Chr(10) & "You can enter Effective Date within 90 Days or contact Corporate for Retro Approval "
                    MsgBox xMsg
                    glbAccessPswd = False
                    frmAccessPswd.Show 1
                    If glbAccessPswd = False Then   'Access Denied
                        'Exit Sub
                        'dlpDate(0).Text = Date
                        dlpDate(0).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    'End If
End If
If glbWFC Then 'Ticket #22964 Franks 12/19/2012
    If WFCMissEndDate4DCPP Then
        MsgBox lblTitle(44).Caption & " is required if Pay Period Amount was changed to zero for DCPP."
        dlpDate(2).SetFocus
        Exit Function
    End If
End If

chkEBENEFITS = True

Exit Function

chkEBENEFITS_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEbenefit", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function chkEBENEFITS1()
Dim SQLQ As String, xok
Dim a As Integer, Msg As String
Dim xlocDays As Integer

chkEBENEFITS1 = False
On Error GoTo chkEBENEFIT1_Err

If Len(clpBCODE.Text) < 1 Then
  MsgBox "Benefit code is a required field", 64
   clpBCODE.SetFocus
  Exit Function
End If
xok = False

Data1.Refresh
If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then      'jaddy 10/20/99
    Data1.Recordset.Find "BF_BCODE='" & IIf(glbLinamar, txtDiv, "") & clpBCODE.Text & "'"
    If Not Data1.Recordset.EOF Then xok = True
    Data1.Refresh
End If
If Not xok Then
  MsgBox "Benefit master not assigned to this employee", 48
   clpBCODE.SetFocus
  Exit Function
End If

If Len(txtBeneName) <= 0 Then
  MsgBox "Beneficiary's Name is Missing", 48
  If txtBeneName.Enabled Then txtBeneName.SetFocus
  Exit Function
End If

If Len(dlpDate(1).Text) > 0 Then
  If Not IsDate(dlpDate(1).Text) Then
    MsgBox "Not a valid date."
    dlpDate(1).SetFocus
    Exit Function
  End If
End If
If Len(MedSplitPc) = 0 Then
    MedSplitPc = 0
End If
If Not IsNumeric(MedSplitPc) Then
    MsgBox "Invalid Percentage"
    MedSplitPc.SetFocus
    Exit Function
End If
If MedSplitPc > 1 Or MedSplitPc < 0 Then
    MsgBox "Invalid Percentage"
    MedSplitPc.SetFocus
    Exit Function
End If

'Ticket #16812
If glbWFC Then
    xBensChg = False 'Ticket #24275 Franks 08/27/2013
    xCurrentBens = True
    xAsOfBens = "" 'Ticket #24275 Franks 08/27/2013
    
    If UCase(comRelation.Text) = "WIFE" Or UCase(comRelation.Text) = "HUSBAND" Then
        MsgBox "'Wife' and 'Husband' are invalid Relationship. Please select 'Spouse'"
        comRelation.SetFocus
        Exit Function
    End If
    If glbEmpCountry = "CANADA" Then
        If UCase(comRelation.Text) = "SPOUSE" Then
            If Len(dlpDate(1).Text) = 0 Then
                MsgBox "Date of Birth is required if Relationship is 'Spouse'"
                dlpDate(1).SetFocus
                Exit Function
            End If
        End If
        
        'Ticket #22009 Franks 05/11/2012
        'Beneficiaries - Can't enter a new SPOUSE record unless the old Spouse has a Date of Separation or Date of Death
        'Ticket #24275 Franks 08/29/2013 - begin
        If clpBCODE.Text = "DB" Then 'Ticket #25028 Franks 01/31/2014
            If fglbNew And Not glbtermopen Then
                If comRelation.Text = "Spouse" Then
                    If isAnotherSameSIN(glbLEE_ID) Then 'Ticket #25026 Franks 02/11/2014
                        'MsgBox "Can't enter a new SPOUSE record unless the old Spouse " & Chr(10) & "has a Date of Separation or Date of Death."
                        'comRelation.SetFocus
                        Exit Function
                    End If
                    If isSpouseExistWithoutEndDate(glbLEE_ID, clpCode(0).Text, "Y") Then
                        MsgBox "Can't enter a new SPOUSE record unless the old Spouse " & Chr(10) & "has a Date of Separation or Date of Death."
                        comRelation.SetFocus
                        Exit Function
                    End If
                   If isSpouseExistWithoutEndDate(glbLEE_ID, clpCode(0).Text, "N") Then
                        MsgBox "There is a current beneficiary with a relationship not equal to Spouse. This beneficiary needs to have an End Date entered before a new beneficiary record."
                        comRelation.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
        'Ticket #24275 Franks 08/29/2013 - end
    End If

    'Ticket #21021 Franks 02/28/2012 - begin
    If Len(dlpDate(3).Text) > 0 Then
        If Not IsDate(dlpDate(3).Text) Then
            MsgBox "Invalid Date"
            dlpDate(3).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpSepDate.Text) > 0 Then
        If Not IsDate(dlpSepDate.Text) Then
            MsgBox "Invalid Date"
            dlpSepDate.SetFocus
            Exit Function
        End If
    End If

    'o   For DB Pension only, if the Beneficiary Name is changed, pop up a message saying "Is this a new Beneficiary".
    'If "Yes", create a new Beneficiary record. Make the "Percentage" on the old Beneficiary say 0%
    'and force the user to enter the Date of Separation, Separation Agreement on File and Spouse Entitled to Pension
    'Ticket #22392 Franks 08/03/2012 added UCase(comRelation.Text) = "SPOUSE"
    If clpBCODE.Text = "DB" And UCase(comRelation.Text) = "SPOUSE" Then  'Ticket #21021
        '''Ticket #24275 Franks 08/27/2013 name can't change, so don't use this logic
        ''If Len(OBNAME) > 0 And Not fglbNew Then 'modify only
        ''    If Not OBNAME = txtBeneName Then
        ''        If Not locSepFlag Then
        ''            Msg = "Beneficiary's Name has been changed." & Chr(10)
        ''            Msg = Msg & "Is this a new Beneficiary?"
        ''
        ''            a% = MsgBox(Msg, 36, "Confirm ")
        ''            If a% <> 6 Then 'NO
        ''                locSepFlag = False
        ''            Else 'Yes
        ''                locSepFlag = True
        ''                If Len(dlpSepDate.Text) = 0 Then
        ''                    Msg = "Please enter Separation Agreement on file (Y/N)," & Chr(10)
        ''                    Msg = Msg & "Spouse Entitled to pension (Y/N) and Date of Separation."
        ''                    MsgBox Msg
        ''                    dlpSepDate.SetFocus
        ''                    Exit Function
        ''                End If
        ''            End If
        ''        Else
        ''            If Len(dlpSepDate.Text) = 0 Then
        ''                MsgBox "Please enter Date of Separation."
        ''                dlpSepDate.SetFocus
        ''                Exit Function
        ''            End If
        ''        End If
        ''    End If
        ''End If
        
        'Ticket #24275 Franks 08/27/2013
        '"   Logic for End Date
        'o   If an End Date is entered and the Relationship equals SPOUSE, display a message saying "Spouse cannot have an End Date. To terminate a spousal beneficiary, a Date of Separation or Date of Death must be entered." Only button is OK.
        'o   Reason for Change must be entered
        'o   Update Pension Beneficiary record. Current will be false and the As of Date equals the End Date.
        If dlpDate(4).Visible Then
            If clpBCODE.Text = "DB" And UCase(comRelation.Text) = "SPOUSE" Then
                If IsDate(dlpDate(4).Text) Then 'End Date
                    If Not ODateEnd = dlpDate(4).Text Then 'entered End Date
                        MsgBox "Spouse cannot have an End Date. To terminate a spousal beneficiary, a Date of Separation or Date of Death must be entered"
                        dlpDate(4).Text = ""
                        'txtBeneName.SetFocus
                        dlpDate(4).SetFocus
                        Exit Function
                    End If
                End If

                If IsDate(dlpSepDate.Text) Then 'Date of Separation
                    If Not ODateSepa = dlpSepDate.Text Then 'entered Date of Separation
                        If chkSpouseEnt.Value Then
                            'o   If a Date of Separation is entered and the Spouse Entitled to Pension is checked, no change to the current flag and as of date in Pension Beneficiary.
                        Else
                            'o   If a Date of Separation is entered and the Spouse Entitled to Pension is not checked, update the Pension Beneficiary's Current to equal false and the As of Date equals the Date of Separation.
                            xBensChg = True
                            xCurrentBens = False
                            xAsOfBens = dlpSepDate.Text
                        End If
                    End If
                End If
                If IsDate(dlpDate(3).Text) Then 'Date of Death
                'o   If Date of Death is entered, update the Pension Beneficiary's Current to equal false and the As of Date equals the Date of Death
                    If Not ODateDeath = dlpDate(3).Text Then 'entered Date of Death
                        xBensChg = True
                        xCurrentBens = False
                        xAsOfBens = dlpDate(3).Text
                    End If
                End If
            End If
        End If
        
    End If
    'Ticket #21021 Franks 02/28/2012 - end
    
    'Ticket #24275 Franks 08/27/2013 - begin
    If clpBCODE.Text = "DB" And Not UCase(comRelation.Text) = "SPOUSE" Then
        'o   If the employee has a DB Beneficiary with an eligible SPOUSE in any Pension Type, the relationship for the new beneficiary is SPOUSE. User cannot change this.
        '"   An eligible spouse is a beneficiary with a Relationship equal to SPOUSE and does not have a Date of Death or Date of Separation.
        If fglbNew Then
            If IsEligibleSpouseExist(glbLEE_ID, txtBeneName.Text) Then
                MsgBox "There is an eligible spouse beneficiary. Cannot add other beneficiary. "
                'txtBeneName.SetFocus
                If txtBeneName.Enabled Then txtBeneName.SetFocus
                Exit Function
            End If
        End If
        If IsDate(dlpDate(4).Text) Then
            If Not ODateEnd = dlpDate(4).Text Then 'entered End Date
                If Len(clpCode(2).Text) = 0 Then
                    MsgBox "Reason for Change is required if End Date is entered."
                    clpCode(2).SetFocus
                    Exit Function
                Else
                    'enter end date and reason
                    xBensChg = True
                    xCurrentBens = False
                    xAsOfBens = dlpDate(4).Text
                End If
            End If
        End If
        'If clpBCODE.Text = "DB" And Not UCase(comRelation.Text) = "SPOUSE" Then
        'Ticket #24830 Franks 01/06/2013  - check on SPOUSE should also include "Common Law".
        If Not clpBCODE.Text = "DB" And Not UCase(comRelation.Text) = "SPOUSE" And Not UCase(comRelation.Text) = "COMMON LAW" Then
            If IsDate(dlpDate(3).Text) Then
                    MsgBox "If a Date of Death is entered, Relationship must equal 'Spouse' or 'Common Law'"
                    dlpDate(3).Text = ""
                    dlpDate(3).SetFocus
                    Exit Function
            End If
            If IsDate(dlpSepDate.Text) Then
                    MsgBox "If a Date of Separation is entered, Relationship must equal 'Spouse' or 'Common Law'"
                    dlpSepDate.Text = ""
                    dlpSepDate.SetFocus
                    Exit Function
            End If
        End If
    End If
    If clpBCODE.Text = "DB" Then
        If Len(clpCode(0).Text) = 0 Then
            MsgBox "Pension Type is required for 'DB' benefit."
            clpCode(0).SetFocus
            Exit Function
        End If
        If InStr(1, xPenTypeList, "'" & clpCode(0).Text & "'") = 0 Then 'Ticket #24337 Franks 09/10/2013
            MsgBox "Pension Type is not valid for this employee." & Chr(10) & "Please select it from Pension Type Lookup."
            clpCode(0).SetFocus
            Exit Function
        End If
        
        'Ticket #24388 Franks 10/16/2013
        If IsDate(dlpDate(3).Text) Then 'Date of Death
            If Not ODateDeath = dlpDate(3).Text Then 'entered End Date
                If Len(dlpSepDate.Text) = 0 And chkSpouseEnt.Value Then
                'o   If only Date of Death is entered, Spouse Entitled to Pension cannot be checked.
                    MsgBox "If only Date of Death is entered, Spouse Entitled to Pension cannot be checked."
                    dlpDate(3).SetFocus '
                    Exit Function
                End If
            End If
            If IsDate(dlpSepDate.Text) Then
                If chkSpouseEnt.Value Then
                    MsgBox "If Dates of Death and Separation are both entered, the Spouse Entitled to Pension cannot be checked"
                    dlpDate(3).SetFocus '
                    Exit Function
                End If
                If CVDate(dlpDate(3).Text) < CVDate(dlpSepDate.Text) Then
                    MsgBox "Date of Death cannot be less than Date of Separation."
                    dlpDate(3).SetFocus '
                    Exit Function
                End If
            End If
        End If
        
        If Len(MedSplitPc.Text) = 0 Then MedSplitPc.Text = 0
        
        'Ticket #24337 Franks 10/01/2013
        'comment out this function because "isSpouseExistWithoutEndDate" doese the same thing
        'Call WFCMultiBeneficiares
    End If
    'Ticket #24275 Franks 08/27/2013 - end
    
    'Ticket #24275 Franks 08/26/2013 - begin
    If dlpDate(4).Visible And Len(dlpDate(4).Text) > 0 Then
        If Not IsDate(dlpDate(4).Text) Then
            MsgBox "Invalid Date"
            dlpDate(4).SetFocus
            Exit Function
        End If
    End If
    If clpCode(2).Visible And clpCode(2).Caption = "Unassigned" Then
        MsgBox "Reason for Change code must be valid", 48
        clpCode(2).SetFocus
        Exit Function
    End If
    If clpCode(0).Visible And clpCode(0).Caption = "Unassigned" Then
        MsgBox "Pension Type code must be valid", 48
        clpCode(0).SetFocus
        Exit Function
    End If
    'Ticket #24275 Franks 08/26/2013 - end
End If


chkEBENEFITS1 = True

Exit Function

chkEBENEFIT1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEbenefit", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub chkSepAgree_Click(Value As Integer)
    If chkSepAgree.Value Then
        chkSpouseEnt.Enabled = True
        lblTitle(50).Enabled = True
    Else
        chkSpouseEnt.Value = False
        chkSpouseEnt.Enabled = False
        lblTitle(50).Enabled = False
    End If
End Sub

Private Sub chkSPlan_Click()
If chkSPlan.Value Then
    lblTitle(41).Visible = True
    lblTitle(42).Visible = True
    lblTitle(43).Visible = True
    txtSComp.Visible = True
    txtSPlan.Visible = True
    txtSCert.Visible = True
Else
    lblTitle(41).Visible = False
    lblTitle(42).Visible = False
    lblTitle(43).Visible = False
    txtSComp.Visible = False
    txtSPlan.Visible = False
    txtSCert.Visible = False
    txtSComp.Text = ""
    txtSPlan.Text = ""
    txtSCert.Text = ""
End If
End Sub

Private Sub chkSPlan_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
oldCode = clpCode(1).Text
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
Dim Msg
Dim xDateAge65
Dim xBenType As String

'If Not cmdOK.Enabled Then Exit Sub
If glbCElgin Then    'glbCompSerial = "S/N - 2292W" Then
    Call CalElgin
End If
If Index = 1 And glbLinamar Then
    Call BENCode_Desc
    If Actn = "A" Or Len(dlpDate(0)) = 0 Or OBCode <> clpCode(1) Then
        txtWaitPeriod = fglbTB_WP
        txtDWM = fglbDWM
        Call txtWaitPeriod_LostFocus
    End If
End If
If Index = 1 Then
    If oldCode <> clpCode(1).Text Then
        If glbWFC And (oldCode = "HCSA" Or oldCode = "HCSA1") Then 'Ticket #22411 Franks 08/08/2012
            Call ResetValues("Y")
        Else
            Call ResetValues
        End If
        
        'Ticket #25500 - Goodmans - Unit Cost/Rate from Benefits Rate table
        'If glbCompSerial = "S/N - 2290W" And (clpCode(1) = "LIFE" Or clpCode(1) = "SLIFE" Or clpCode(1) = "CLIFE" Or clpCode(1) = "OLIFE") Then
        'If glbCompSerial = "S/N - 2290W" And (clpCode(1) = "SLIFE" Or clpCode(1) = "OLIFE") Then
        'Ticket #27113 - Making option to have different types of Benefit Code setup under Benefit Rates table
        xBenType = ""
        xBenType = Get_BenefitType_BenefitRateTable(clpCode(1).Text)
        If glbCompSerial = "S/N - 2290W" Then
            'If Left(clpCode(1).Text, 1) = "S" Then
            '    medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, Spouse)
            'ElseIf Left(clpCode(1).Text, 1) = "C" Then
            '    medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, Children)
            'ElseIf Left(clpCode(1).Text, 1) = "O" Then
            '    medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, DependentRelationship.Employee)
            'End If
            
            If xBenType = "S" Then
                medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, Spouse)
            ElseIf xBenType = "O" Then
                medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, Children)
            ElseIf xBenType = "E" Then
                medUnitCost = Get_BenefitRate(glbLEE_ID, clpCode(1).Text, DependentRelationship.Employee)
            End If
            
            Call Set_SalCover
        End If
        
        'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
        If glbCompSerial = "S/N - 2290W" And (clpCode(1) = "LTD" And ((clpGroup <> "PARTNERS" And clpGroup <> "ART") Or (clpGroup = ""))) Then 'And cmbDWM.ListIndex >= 0 Then
            'xPER = Left(cmbDWM, 1)
            'If xPER = "W" Then xPER = "ww"
            
            'Get the date for Age 65 or 67 based on the Benefit Group
            If (clpGroup = "") Then
                xDateAge65 = DateAdd("yyyy", 67, CVDate(GetEmpData(glbLEE_ID, "ED_DOB")))
            Else
                xDateAge65 = DateAdd("yyyy", 65, CVDate(GetEmpData(glbLEE_ID, "ED_DOB")))
            End If
            
            'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
            'dlpDate(2).Text = MonthLastDate(DateAdd(xPER, 0 - Val(txtWaitPeriod), CVDate(xDateAge65)))
            'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
            If (clpGroup = "") Then
                dlpDate(2).Text = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
            Else
                dlpDate(2).Text = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
            End If
        End If
    End If
End If
End Sub

Private Sub clpGroup_GotFocus()
OldGroup = clpGroup.Text
End Sub

Private Sub clpGroup_LostFocus()
If fEBGroup <> "NOGROUP" And (Trim(clpGroup.Text) = "" Or clpGroup.Caption <> "Unassigned") Then
    Call Set_Group_Benefit
End If
If OldGroup <> clpGroup.Text Then
    Call ResetValues
End If
End Sub

Private Sub cmbDWM_Click()
If Left(cmbDWM, 1) <> txtDWM Then
    Call txtWaitPeriod_LostFocus
End If
End Sub

Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbDWM_LostFocus()
'Call txtWaitPeriod_LostFocus
End Sub

Private Sub cmbPerOrDoll_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdBens_Click()

If FrmDetails.Visible = True Then
    FrmDetails.Visible = False
    cmdBens.Caption = "Benefits"
    glbSkip = True
    Data1.Refresh
    Dim xCodeList As String
    Do Until Data1.Recordset.EOF
        xCodeList = xCodeList & "," & Data1.Recordset!BF_BCODE
        Data1.Recordset.MoveNext
    Loop
    If Len(xCodeList) > 0 Then xCodeList = Mid(xCodeList, 2)
    clpBCODE.seleEMPCode = xCodeList 'Pay attention
    Data1.Refresh
    Data3.Refresh
    vbxTrueGrid.Visible = False
    TblBENS.Visible = True
    FrmBens.Visible = True
    FrmBens.Enabled = False
    cmdRecal.Visible = False
    cmdRecalAll.Visible = False
    fraCopyBS.Visible = False
    cmdBsCoopy.Visible = True
    cmdBsCoopy.Enabled = True
    MDIMain.panHelp(0).Caption = "Listing of Beneficiary "
    scrControl.Visible = False
    
    If Not gSec_Upd_Beneficiary Then
        cmdBsCoopy.Enabled = False
    End If
    
    Call Display_Value
    
    If glbWFC Then 'Ticket #24275 Franks 08/28/2013
        If FrmBens.Visible Then
            If WFCisEmpNoChanged Then
                Call WFCPenTypeList
                If Not WFCIsAllPenTypeCurrent Then
                    MsgBox "This employee does not have a beneficiary assigned to all pension types. Please add the beneficiary for the missing Pension Type(s)."
                End If
            End If
            Call WFCPenFieldsEnable(False) 'Ticket #24317 Franks 09/17/2013
            If glbtermopen Then locWFCEmpID = glbTERM_Seq Else locWFCEmpID = glbLEE_ID
        End If
    End If
Else
    glbSkip = False
    cmdBens.Caption = "Beneficiary"
    Data1.Refresh
    clpBCODE.seleEMPCode = ""
    vbxTrueGrid.Visible = True
    FrmDetails.Visible = True
    TblBENS.Visible = False
    FrmBens.Visible = False
    cmdRecal.Visible = True
    cmdRecalAll.Visible = True
    cmdBsCoopy.Visible = False
    fraCopyBS.Visible = False
    MDIMain.panHelp(0).Caption = "Listing of Benefits "
    Call Form_Resize
End If

End Sub

Private Sub cmdBSCancel_Click()
fraCopyBS.Visible = False
cmdBsCoopy.Enabled = True
End Sub

Private Sub cmdBsCoopy_Click()
    fraCopyBS.Visible = True
    cmdBsCoopy.Enabled = False
    'clpBCODF.Text = clpBCODE.Text
    If glbWFC Then 'Ticket #24275 Franks 08/28/2013
        lblTitle(55).Visible = True
        clpPenType.Visible = True
        clpBCODT.Text = ""
        clpPenType = ""
    End If
End Sub

Sub cmdCancel_Click()
Dim x, bk

On Error GoTo Can_Err

fglbNew = False

DoEvents

Call Display_Value
Call getCodes 'Ticket #24337 Franks 09/30/2013
Call WFCPenFieldsEnable(False) 'Ticket #24337 Franks 09/30/2013

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRBENFT", "Cancel")
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
If glbOnTop = "FRMEBENEFITS" Then glbOnTop = ""
End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x%
Dim xBCode As String
Dim xBEffDate


If FrmDetails.Visible = True Then
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
Else
    If Data3.Recordset.BOF And Data3.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    If glbWFC Then
        If clpBCODE.Text = "DB" Then
            'MsgBox "Can not delete DB Beneficiary recrord"
            'Exit Sub
            'Ticket #24275 Franks 08/27/2013
            glbAccessPswd = False
            frmAccessPswd.Show 1
            If glbAccessPswd = False Then   'Access Denied
                Exit Sub
            End If
        End If
    End If
End If
On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

If FrmDetails.Visible = True Then

    'Release 8.1 - To be used when sending email about the Delete
    xBCode = Data1.Recordset("BF_BCODE")
    xBEffDate = Data1.Recordset("BF_EDATE")
    
    x% = Delete_Match_BENS()
    
    If Not glbtermopen Then
        If Not AUDITBENF("D", 1) Then MsgBox "ERROR - AUDIT FILE"
    End If
    If glbtermopen Then
      gdbAdoIhr001X.BeginTrans
    
      rsDATA.Delete
      gdbAdoIhr001X.CommitTrans
      Data1.Refresh
    Else
      gdbAdoIhr001.BeginTrans
      rsDATA("BF_LUSER") = glbUserID
      rsDATA.Update
      rsDATA.Delete
      gdbAdoIhr001.CommitTrans
      Data1.Refresh
    End If
        
    'Release 8.1 - Send email on Benefit changes
    If gsEMAIL_ONBENEFIT Then
        'If glbCompSerial = "S/N - 2382W" Then  'Samuel
        '    MailBody = GetEmailBodyForSamuel(glbLEE_ID)
        '    MailBody = MailBody & "has existing Benefit '" & GetTABLDesc("BNCD", xBCode) & "' Deleted "
        '    MailBody = MailBody & "with Effective Date " & Format(CVDate(xBEffDate), "SHORT DATE") & vbCrLf
        '    Call EmailSendingForSamuel
        'Else
            MailBody = "The Deleted Benefit:" & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            MailBody = MailBody & "Deleted Benefit: " & GetTABLDesc("BNCD", xBCode) & vbCrLf
            MailBody = MailBody & "Effective Date: " & Format(CVDate(xBEffDate), "SHORT DATE") & vbCrLf
            Call imgEmail_Click("DELETE")
        'End If
        Screen.MousePointer = DEFAULT
    End If
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Call Display_Value
    End If
    fglbNew = False
    Call SET_UP_MODE

Else
    If glbWFC Then 'Ticket #24275 Franks 08/27/2013
      If clpBCODE.Text = "DB" Then 'delete the pension beneficiary too.
          Call WFCDelPenBeneficiary(glbLEE_ID, txtBeneName.Text, clpCode(0).Text)
      End If
    End If
    If Not glbtermopen Then
        If Not AUDITBENF("D", 2) Then MsgBox "ERROR - AUDIT FILE"
    End If
    If glbtermopen Then
      gdbAdoIhr001X.BeginTrans
      rsDATA3.Delete
      gdbAdoIhr001X.CommitTrans
      Data3.Refresh
    Else
      gdbAdoIhr001.BeginTrans
      rsDATA3.Delete
      gdbAdoIhr001.CommitTrans
      Data3.Refresh
    End If
    If Data3.Recordset.EOF And Data3.Recordset.BOF Then
        Call Display_Value
    End If
    Call SET_UP_MODE2

End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRBENFT", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err
IfElginLife = False
Actn = "M"
If FrmDetails.Visible = True Then
  OBCode = clpCode(1).Text
  OCOVER = txtCovType
  OTCOST = medTCost
  OPremium = lblAP
  OPPE = medPPE
  OPCC = medPPComp
  OPPAMT = medPayPeriodAmount
  OTOTAL = medTCost
  OMAXDOL = medMaxAmnt
  OEDate = dlpDate(0).Text
  OPER = txtPer
  OMTHCOMP = medMCCOST 'ADDED BY RAUBREY 7/9/97
  OMTHEMP = medMECOST 'ADDED BY RAUBREY 7/9/97
  OTAXBEN = txtTAXBEN 'ADDED BY RAUBREY 7/9/97
  OBAMT = medCovAmount
  OUNITCOST = medUnitCost
  OBenEndDate = dlpDate(2).Text
 ' clpCode(1).SetFocus
Else
  OBCode = clpBCODE.Text
  OBNAME = txtBeneName
  OBRELATE = lblRel
  ODOB = dlpDate(1).Text
  ODateEnd = dlpDate(4).Text 'Ticket #24275 Franks 08/27/2013
  ODateDeath = dlpDate(3).Text
  ODateSepa = dlpSepDate.Text
  xCurrentBens = False
  xAsOfBens = "" 'Ticket #24275 Franks 08/27/2013
  OMedSplitPc = MedSplitPc.Text
  locSepFlag = False
End If
Call SET_UP_MODE
'vbxTrueGrid.Enabled = False
'FrmDetails.Enabled = True
If glbCompSerial = "S/N - 2214W" And optActual(0) Then
    medMCCOST.Enabled = True
    medMECOST.Enabled = True
End If
Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub CaseyHouseBenefit()
    If Val(medMCCOST) > 0 Or Val(medMECOST) > 0 Then
        medPPComp = Val(medMCCOST) / (Val(medMCCOST) + Val(medMECOST))
        medPPE = Val(medMECOST) / (Val(medMCCOST) + Val(medMECOST))
    End If
    medCompCost = Val(medMCCOST) * 12
    medEECost = Val(medMECOST) * 12
    medTCost = Val(medCompCost) + Val(medEECost)
    'FlagCaseyHouse = True
End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True

Call SET_UP_MODE

On Error GoTo AddN_Err

Actn = "A"
If FrmDetails.Visible = True Then
    Call Set_Control("B", Me)
    OBCode = ""
    OCOVER = ""
    OTCOST = ""
    OPremium = ""
    OPPE = ""
    OPCC = ""
    OPPAMT = ""
    OMAXDOL = ""
    OEDate = ""
    OPER = ""
    OBAMT = ""
    OTOTAL = ""
    OUNITCOST = ""
    OBenEndDate = ""
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    lblCNum.Caption = "001"
    medPPComp = 1
    medPPE = 0
    medTCost = 0
    medUnitCost = 0
    txtPer = 0
    comSalDepn = "No"
    If glbCElgin Then txtDWM = "D"

    If clpGroup.Enabled And clpGroup.Visible Then
        clpGroup.SetFocus
    Else
        clpCode(1).SetFocus
    End If
    If glbCompSerial = "S/N - 2214W" And optActual(0) Then
        medMCCOST.Enabled = True
        medMECOST.Enabled = True
    End If
    clpCode(1) = ""
Else
    Call Set_Control3("B", rsDATA3)
    clpBCODE = ""
    'WFC Pension Outstanding Tasks By Dec2109.doc
    lblRel.Caption = "Spouse"
    Call WFCPenFieldsEnable(True) 'Ticket #24275 Franks 08/27/2013
    If lblRel.Caption = "Spouse" Then Call WFCEligibleSpouseFields 'Ticket #24275 Franks 08/29/2013
End If

Exit Sub
AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRBENFT", "Add")
Resume Next
End Sub

   
'Private Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim rsBen As New ADODB.Recordset
Dim xID
Dim a As Integer, Msg As String, x%
Dim locBCode, locBenEndDate
Dim xNewUpt As String 'Ticket #22884
Dim isNewGPBeneFunc As Boolean

On Error GoTo Add_Err

'Call MedSplitPc_LostFocus


If FrmBens.Visible Then Call comRelation_LostFocus

If glbWFC Then
    If glbPlantCode = "GREN" Then
        If Val(medPayPeriodAmount.Text) = 0 And Val(medTCost.Text) > 0 Then '
            If GetPayPeriod(lblEENum) = "W" Then
                medPayPeriodAmount.Text = Round(medTCost.Text / 52, 2)
            Else
                medPayPeriodAmount.Text = Round(medTCost.Text / 12, 2)
            End If
        End If
    Else
        ' danielk - 03/19/2003 - Pay Period Amount should default to Employee Monthly / 4.33
        If Val(medPayPeriodAmount.Text) = 0 And Val(medMECOST.Text) > 0 Then 'GetPayPeriod
            If GetSalCD(lblEENum) = "H" Then
                If IsDate(dlpDate(2).Text) Then 'Ticket #30446 Franks 08/09/2017
                    '"   When I zero out Pay Period Amount, it doesn't save the zero. The old dollar amount is put back into the field.
                    'If the benefits have an END DATE, this shouldn't happen.
                Else
                    medPayPeriodAmount.Text = medMECOST.Text / 4.33
                End If
            End If
        End If
    End If
End If
If FrmDetails.Visible = True Then
    If Not glbCElgin Then
        Call Set_SalCover
    Else
        Call CalElgin
    End If
    
    If Not chkEBENEFITS() Then Exit Sub
    
    locBCode = clpCode(1).Text: locBenEndDate = dlpDate(2).Text
    rsDATA.Requery
    If fglbNew Then
        If glbCompSerial = "S/N - 2387W" Then 'Bird Packaging Limited 'Ticket #13701
            If clpCode(1).Text = "PEN" Then
                medPayPeriodAmount = GetPayrollPension
            End If
        End If
        If gsEMAIL_ONBENEFIT Then
            If NewHireForms.count = 0 Then 'Non new hire
                If glbCompSerial = "S/N - 2382W" Then  'Samuel
                    MailBody = GetEmailBodyForSamuel(glbLEE_ID)
                    MailBody = MailBody & "has New Benefit " & GetTABLDesc("BNCD", clpCode(1)) & " "
                    MailBody = MailBody & "on Effective Date " & dlpDate(0) & vbCrLf
                    Call EmailSendingForSamuel
                Else
                    MailBody = "The New Benefit:" & vbCrLf & vbCrLf
                    MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                    MailBody = MailBody & "New Benefit: " & GetTABLDesc("BNCD", clpCode(1)) & vbCrLf
                    MailBody = MailBody & "Effective Date: " & dlpDate(0) & vbCrLf
                    Call imgEmail_Click
                End If
                Screen.MousePointer = DEFAULT
                
                'Screen.MousePointer = HOURGLASS
            End If
        End If
        
        'Ticket #19255
        If glbLambton And glbVadim Then 'Ticket #25931
            txtPayrollID.Text = Get_Payroll_ID_For_Benefit(lblEEID)
        Else
            txtPayrollID.Text = GetEmpData(lblEEID, "ED_PAYROLL_ID")
        End If
        
        rsDATA.AddNew
    Else
        'Release 8.1 - Send Email Notification on Benefit change as well
        If gsEMAIL_ONBENEFIT Then
            If OBCode <> clpCode(1).Text Or OCOVER <> txtCovType Or OTCOST <> medTCost Or OPremium <> lblAP Or OPPE <> medPPE Or _
                OPCC <> medPPComp Or OPPAMT <> medPayPeriodAmount Or OTOTAL <> medTCost Or OMAXDOL <> medMaxAmnt Or _
                OEDate <> dlpDate(0).Text Or OPER <> txtPer Or OMTHCOMP <> medMCCOST Or OMTHEMP <> medMECOST Or OTAXBEN <> txtTAXBEN Or _
                OBAMT <> medCovAmount Or OUNITCOST <> medUnitCost Or OBenEndDate <> dlpDate(2).Text Then
                
                MailBody = "The Updated Benefit:" & vbCrLf & vbCrLf
                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                MailBody = MailBody & "Updated Benefit: " & GetTABLDesc("BNCD", clpCode(1)) & vbCrLf
                MailBody = MailBody & "Effective Date: " & dlpDate(0) & vbCrLf
                
                Call imgEmail_Click("UPDATE")
            End If
            Screen.MousePointer = DEFAULT
        End If
    End If
    
    'Ticket #19255
    If Len(Trim(txtPayrollID)) = 0 Then
        If glbLambton And glbVadim Then 'Ticket #25931
            txtPayrollID.Text = Get_Payroll_ID_For_Benefit(lblEEID)
        Else
            txtPayrollID.Text = GetEmpData(lblEEID, "ED_PAYROLL_ID")
        End If
    End If
    
    If Not glbtermopen Then
        If Not AUDITBENF(Actn, 1) Then MsgBox "ERROR - AUDIT FILE"
        If glbWFC Then
            Call AUDIT_MANULIFE_BENF(clpCode(1).Text, dlpDate(0).Text, txtCovType.Text, txtPolicy.Text, dlpDate(2).Text)
        End If
    Else
        rsDATA!TERM_SEQ = glbTERM_Seq
    End If
    
    Call UpdUStats(Me)
    
    If glbVadim Then
        If fglbNew Then
            Updstats(0) = dlpDate(0)
        Else
            Updstats(0) = getProcessDate
        End If
    End If
    
    If IsNumeric(txtWaitPeriod) Then Call txtWaitPeriod_LostFocus
    
    txtDWM = Left(cmbDWM, 1)
    txtPerOrDoll = Left(cmbPerOrDoll, 1)
    
    Call UpdCodes1
    
    If glbtermopen Then
        gdbAdoIhr001X.BeginTrans
        Call Set_Control("U", Me, rsDATA)
        rsDATA.Update
        gdbAdoIhr001X.CommitTrans
        rsDATA.Resync
        xID = rsDATA!BF_BENE_ID
        Data1.Refresh
    Else
        gdbAdoIhr001.BeginTrans
        Call Set_Control("U", Me, rsDATA)
        rsDATA.Update
        gdbAdoIhr001.CommitTrans
        'Apr 17,2003, comment by Frank Ticket# 4032
        '"Insufficient Key Column Information for updating and refreshing"
        'rsDATA.Resync
        xID = rsDATA!BF_BENE_ID
        Data1.Refresh
    End If
    
    If glbWFC Then
        If OBenEndDate = "" And IsDate(locBenEndDate) Then
            Msg = "Do you want this Benefit End Date to be used for all benefits "
            a% = MsgBox(Msg, 36, "Confirm Update")
            If a% = 6 Then 'Exit Sub
                Call UpdBeneEndDate(locBenEndDate, "UPD_END_DATE", "")
                Data1.Refresh
            End If
        End If
        If IsDate(OBenEndDate) And locBenEndDate = "" Then
            'If the employee returns before the Benefit End Date,
            'the HR Administrator will need to go into the Benefits screen and delete the Benefit End Date.
            'This change will delete the record from the Manulife Transaction Audit Table.
            Call UpdBeneEndDate(OBenEndDate, "DEL_END_DATE", locBCode)
        End If
    End If
    
    Data1.Recordset.Find "BF_BENE_ID=" & xID

Else
    Call comRelation_LostFocus
    If Not chkEBENEFITS1() Then Exit Sub
    rsDATA3.Requery
    If fglbNew Then rsDATA3.AddNew
    
    If Not glbtermopen Then
        If Not AUDITBENF(Actn, 2) Then MsgBox "ERROR - AUDIT FILE"
    Else
        rsDATA3!TERM_SEQ = glbTERM_Seq
    End If
    Call UpdCodes2
'    'rsDATA3.AddNew
    rsDATA3!BD_COMPNO = "001"

    If glbtermopen Then
       rsDATA3!TERM_SEQ = glbTERM_Seq
    End If
    rsDATA3!BD_EMPNBR = glbLEE_ID
    rsDATA3!BD_LDATE = Format(Now, "SHORT DATE")
    rsDATA3!BD_LTIME = Time$
    rsDATA3!BD_LUSER = glbUserID
    If glbWFC Then
        'Ticket #21021 Franks 02/28/2012 - begin
        'Keep the old DB Beneficiaries Name, the new name will be a new record
        NewBNAME = txtBeneName.Text
        ''Ticket #24275 Franks 08/27/2013
        'If locSepFlag Then
        '    rsDATA3!BD_BNAME = OBNAME
        '    txtBeneName.Text = OBNAME
        '    rsDATA3!BD_PC = 0
        '    MedSplitPc.Text = 0
        'End If
    End If
    If glbtermopen Then
        gdbAdoIhr001X.BeginTrans
        Call Set_Control3("U", rsDATA3)
        rsDATA3.Update
'        rsDATA3.Requery
        xID = rsDATA3!BD_ID
        gdbAdoIhr001X.CommitTrans
        Data3.Refresh
    Else
        gdbAdoIhr001.BeginTrans
        Call Set_Control3("U", rsDATA3)
        rsDATA3.Update
'        rsDATA3.Requery
        xID = rsDATA3!BD_ID
        gdbAdoIhr001.CommitTrans
        Data3.Refresh
    End If
    Data3.Recordset.Find "BD_ID=" & xID
    Call SET_UP_MODE
    TblBENS.SetFocus
End If

If glbWFC Then
    'Pension Beneficiary - Ticket #16395
    If Not FrmDetails.Visible Then
        '''If fglbNew Then
        ''    'If Left(clpBCODE.Text, 3) = "LIF" Then
        ''    If clpBCODE.Text = "DB" Then
        ''        Call WFCPensionBeneficiary(glbLEE_ID, clpBCODE.Text)
        ''    End If
        '''End If
        
        'Ticket #21021 Franks 02/28/2012 - begin
        If clpBCODE.Text = "DB" Then
            ''Ticket #24275 Franks 08/27/2013
            ''If locSepFlag Then 'User's answer is YES
            ''    'create a new Beneficiary record
            ''    If glbtermopen Then 'Ticket #22884 Franks
            ''        Call WFCNewDBBenforSeparation(glbTERM_ID, clpBCODE.Text, NewBNAME, ODOB, OMedSplitPc, OBRELATE, glbTERM_Seq)
            ''    Else
            ''        Call WFCNewDBBenforSeparation(glbLEE_ID, clpBCODE.Text, NewBNAME, ODOB, OMedSplitPc, OBRELATE)
            ''    End If
            ''    locSepFlag = False
            ''    Data3.Refresh
            ''    Data3.Recordset.Find "BD_ID=" & xID
            ''    Call SET_UP_MODE
            ''End If
            If clpBCODE.Text = "DB" Then 'Ticket #16395
                ''Ticket #24275 Franks 08/27/2013 - begin
                ''If locSepFlag Then xNewUpt = "NEW" Else xNewUpt = "UPT" 'Ticket #22884
                xNewUpt = "NEW" 'Ticket #24275 Franks 08/27/2013
                If glbtermopen Then
                    Call WFCPensionBeneficiary(glbTERM_ID, clpBCODE.Text, glbTERM_Seq, xNewUpt, xBensChg, xCurrentBens, xAsOfBens, clpCode(0).Text, txtBeneName.Text)
                Else
                    Call WFCPensionBeneficiary(glbLEE_ID, clpBCODE.Text, , xNewUpt, xBensChg, xCurrentBens, xAsOfBens, clpCode(0).Text, txtBeneName.Text)
                End If
                If fglbNew And Not UCase(comRelation.Text) = "SPOUSE" Then
                    If Len(xPenTypeList) > 0 Then
                        If InStr(1, xPenTypeList, ",") > 0 Then
                            Msg = "Will this be the beneficiary for all Pension Types? "
                            a% = MsgBox(Msg, 36, "Confirm ")
                            If a% = 6 Then
                                Call WFCMultiPenType(clpCode(0).Text, txtBeneName.Text)
                                Data3.Refresh
                                Data3.Recordset.Find "BD_ID=" & xID
                                Call SET_UP_MODE
                            End If
                        End If
                    End If
                End If
                'Ticket #24275 Franks 08/27/2013 - end
            End If
        End If
        'Ticket #21021 Franks 02/28/2012 - end
    End If
    'Pension DC
    If FrmDetails.Visible Then
        If fglbNew Then
            If (clpCode(1).Text) = "DC" Then
                Call WFCPensionMasOnType(glbLEE_ID, clpCode(1).Text, dlpDate(0).Text)
            End If
        End If
        'Ticket #22964 Franks 12/17/2012
        If (clpCode(1).Text) = "DCPP" Then
            Call WFCPensionMasOnType(glbLEE_ID, clpCode(1).Text, dlpDate(0).Text)
        End If
    End If

End If
fglbNew = False

If glbMediPay Then Call Employee_Benefit_Integration(glbLEE_ID) 'Ticket #14752

If glbGP Then 'Ticket #26654 Franks 02/10/2015
    'isNewGPBeneFunc = False
    'If glbCompSerial = "S/N - 2453W" Then  'Town of Gander
    '    'Call Employee_GP_NewBenefitDeduction_Integration(glbLEE_ID)
    '    isNewGPBeneFunc = True
    'End If
    'If glbCompSerial = "S/N - 2486W" Then isNewGPBeneFunc = True 'Ticket #28995 Franks 01/04/2017
    'If glbCompSerial = "S/N - 2484W" Then isNewGPBeneFunc = True 'Ticket #28396 Franks 03/08/2017
    'If glbCompSerial = "S/N - 2487W" Then isNewGPBeneFunc = True 'Ticket #30217 Franks 06/13/2017
    'If isNewGPBeneFunc Then
    '    Call Employee_GP_NewBenefitDeduction_Integration(glbLEE_ID)
    'End If
    'Ticket #30111 Franks 06/13/2017 - dont use isNewGPBeneFunc here, will check it in GP_Integraion function of "NewBenDeduction"
    Call Employee_GP_NewBenefitDeduction_Integration(glbLEE_ID)
End If

If NextFormIF("Benefit") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRBENFT", "Update")
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
Dim xReport
'~~~ added a second crystal control (array) by RAUBREY 7/7/97
If InStr(cmdBens.Caption, "ficiary") > 0 Then
    RHeading = lblEEName & "'s Benefits Report"
    If glbtermopen Then
       Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenef2.rpt"
       Me.vbxCrystal(0).SelectionFormula = "{Term_HRBENFT.Term_Seq}=" & glbTERM_Seq & " "
    Else
       Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenefi.rpt"
       Me.vbxCrystal(0).SelectionFormula = "{HRBENFT.BF_EMPNBR}=" & glbLEE_ID & " "
    End If
Else
    RHeading = lblEEName & "'s Beneficiaries Report"
    If glbtermopen Then
        Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenes2.rpt"
        Me.vbxCrystal(0).SelectionFormula = "{Term_HRBENS.Term_Seq}=" & glbTERM_Seq & " "
    Else
        Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenefc.rpt"
        Me.vbxCrystal(0).SelectionFormula = "{HRBENS.BD_EMPNBR}=" & glbLEE_ID & " "
    End If
End If

Me.vbxCrystal(1).WindowTitle = RHeading
Me.vbxCrystal(0).Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"

If glbSQL Or glbOracle Then
    Me.vbxCrystal(0).Connect = RptODBC_SQL
Else
    Me.vbxCrystal(0).Connect = "PWD=petman;"
    Me.vbxCrystal(0).DataFiles(0) = IIf(glbtermopen, glbIHRAUDIT, glbIHRDB)
End If

Me.vbxCrystal(0).Destination = 1
Me.vbxCrystal(0).Action = 1
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub

Sub cmdView_Click()
Dim RHeading As String
Dim xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal(0).WindowShowPrintSetupBtn = True
Me.vbxCrystal(1).WindowShowPrintSetupBtn = True

'~~~ added a second crystal control (array) by RAUBREY 7/7/97
If InStr(cmdBens.Caption, "ficiary") > 0 Then
    RHeading = lblEEName & "'s Benefits Report"
    If glbtermopen Then
       Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenef2.rpt"
       Me.vbxCrystal(0).SelectionFormula = "{Term_HRBENFT.Term_Seq}=" & glbTERM_Seq & " "
    Else
       Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenefi.rpt"
       Me.vbxCrystal(0).SelectionFormula = "{HRBENFT.BF_EMPNBR}=" & glbLEE_ID & " "
    End If
Else
    RHeading = lblEEName & "'s Beneficiaries Report"
    If glbtermopen Then
        Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenes2.rpt"
        Me.vbxCrystal(0).SelectionFormula = "{Term_HRBENS.Term_Seq}=" & glbTERM_Seq & " "
    Else
        Me.vbxCrystal(0).ReportFileName = glbIHRREPORTS & "rgbenefc.rpt"
        Me.vbxCrystal(0).SelectionFormula = "{HRBENS.BD_EMPNBR}=" & glbLEE_ID & " "
    End If
End If

Me.vbxCrystal(1).WindowTitle = RHeading
Me.vbxCrystal(0).Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"

If glbSQL Or glbOracle Then
    Me.vbxCrystal(0).Connect = RptODBC_SQL
Else
    Me.vbxCrystal(0).Connect = "PWD=petman;"
    Me.vbxCrystal(0).DataFiles(0) = IIf(glbtermopen, glbIHRAUDIT, glbIHRDB)
End If
Me.vbxCrystal(0).Destination = 0
Me.vbxCrystal(0).Action = 1
End Sub


Private Sub cmdBSOK_Click()
Dim rsTBens As New ADODB.Recordset
Dim SQLQ As String
Dim Msg As String, a%
Dim xToWFCPen As Boolean 'Ticket #24275 Franks 08/28/2013

    If Len(clpBCODE.Text) = 0 Then
        MsgBox "No Beneficiary Record."
        Exit Sub
    End If
    If Len(clpBCODT.Text) = 0 Then
        MsgBox "No To Benefit."
        clpBCODT.SetFocus
        Exit Sub
    End If
    'Ticket #24275 Franks 08/28/2013 - begin
    If glbWFC Then
        xToWFCPen = False
        If WFCIsPenEligilbe Then
            If Len(clpPenType.Text) = 0 And clpBCODT.Text = "DB" Then
                MsgBox "Pension Type is required for Pension Eligible employee and 'DB' Benefit."
                Exit Sub
            End If
            If clpPenType.Caption = "Unassigned" Then
                MsgBox "Invalid Pension Type."
                clpPenType.SetFocus
                Exit Sub
            End If
            'Ticket #24451 Franks 10/08/2013
            If clpBCODT.Text = "DB" Then 'Ticket #25026 Franks 02/06/2014
                If InStr(1, xPenTypeList, "'" & clpPenType.Text & "'") = 0 Then 'Ticket #24337 Franks 09/10/2013
                    MsgBox "Pension Type is not valid for this employee." & Chr(10) & "Please select it from Pension Type Lookup."
                    clpPenType.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If (clpBCODE.Text) = (clpBCODT.Text) Then
            If (clpPenType.Text) = (clpCode(0).Text) Then
                MsgBox "Can not copy to the same benefit code and Pension Type code."
                Exit Sub
            End If
        End If
        Msg = ""
        Msg = Msg & Chr(10) & "Are You Sure You Want To Copy Benefit from '" & clpBCODE.Text & "'"  '& "' to '" & clpBCODT.Text & "'? "
        If Len(clpCode(0).Text) > 0 Then Msg = Msg & " and '" & clpCode(0).Text & "' "
        Msg = Msg & " to '" & clpBCODT.Text & "' "
        If Len(clpPenType.Text) > 0 Then Msg = Msg & " and '" & clpPenType.Text & "' "
        Msg = Msg & "?"
    Else
        Msg = ""
        Msg = Msg & Chr(10) & "Are You Sure You Want To Copy Benefit from '" & clpBCODE.Text & "' to '" & clpBCODT.Text & "'? "
        If (clpBCODE.Text) = (clpBCODT.Text) Then
            MsgBox "Can not copy to the same benefit code."
            Exit Sub
        End If
    End If
    'Ticket #24275 Franks 08/28/2013 - end

    
    Msg = ""
    Msg = Msg & Chr(10) & "Are You Sure You Want To Copy Benefit from '" & clpBCODE.Text & "' to '" & clpBCODT.Text & "'? "
      
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then
        Exit Sub
    End If


    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HRBENS "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND BD_BCODE = '" & clpBCODT.Text & "' "
    Else
        SQLQ = " SELECT * FROM HRBENS "
        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND BD_BCODE = '" & clpBCODT.Text & "' "
    End If
    If glbWFC Then 'Ticket #24275 Franks 08/28/2013
        If Len(clpPenType.Text) > 0 Then
            SQLQ = SQLQ & " AND BD_PENSIONTYPE = '" & clpPenType.Text & "' "
        End If
        'Ticket #25026 Franks 02/06/2014
        If Len(txtBeneName.Text) > 0 Then
            SQLQ = SQLQ & " AND BD_BNAME = '" & txtBeneName.Text & "' "
        End If
    End If
    If rsTBens.State <> 0 Then rsTBens.Close
    rsTBens.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTBens.EOF Then
        rsTBens.AddNew
        rsTBens("BD_COMPNO") = "001"
        If glbtermopen Then
            rsTBens("BD_EMPNBR") = glbTERM_ID
        Else
            rsTBens("BD_EMPNBR") = glbLEE_ID
        End If
        If glbWFC And WFCIsPenEligilbe Then 'Ticket #24275 Franks 08/28/2013
            xToWFCPen = True
        End If
    End If
    rsTBens("BD_BCODE") = clpBCODT.Text
    rsTBens("BD_BNAME") = txtBeneName.Text
    rsTBens("BD_RELATE") = comRelation.Text
    If IsDate(dlpDate(1).Text) Then rsTBens("BD_DOB") = CVDate(dlpDate(1).Text)
    If IsNumeric(MedSplitPc.Text) Then rsTBens("BD_PC") = Val(MedSplitPc.Text)
    rsTBens("BD_LDATE") = Date
    rsTBens("BD_LTIME") = Time$
    rsTBens("BD_LUSER") = glbUserID
    If glbtermopen Then
        rsTBens("TERM_SEQ") = glbTERM_Seq
    End If
    If glbWFC Then 'Ticket #24275 Franks 08/28/2013
        If Len(clpPenType.Text) > 0 Then
            rsTBens("BD_PENSIONTYPE") = clpPenType.Text
        End If
    End If
    rsTBens.Update
    rsTBens.Close
    
    'Ticket #24275 Franks 08/28/2013
    If glbWFC And xToWFCPen Then
        If glbtermopen Then
            Call WFCPensionBeneficiary(glbTERM_ID, "DB", glbTERM_Seq, "NEW", , , , clpPenType.Text, txtBeneName.Text)
        Else
            Call WFCPensionBeneficiary(glbLEE_ID, "DB", , "NEW", , , , clpPenType.Text, txtBeneName.Text)
        End If
    End If
    
    fraCopyBS.Visible = False
    cmdBsCoopy.Enabled = True

    Data3.Refresh

    
End Sub

Private Sub cmdRecal_Click()
Dim bk
Dim TermOrActive, BCodeCover

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bk = Data1.Recordset.Bookmark
End If

Call Recal_Screen_Values

If Not IsNull(Data1.Recordset!BF_GROUP) Then
    If Len(Data1.Recordset!BF_GROUP) > 0 Then
        If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
            If glbtermopen Then TermOrActive = "T" Else TermOrActive = "A"
            BCodeCover = Data1.Recordset!BF_BCODE
            If IsNull(Data1.Recordset!BF_COVER) Then
                BCodeCover = BCodeCover & "_"
            Else
                BCodeCover = BCodeCover & "_" & Data1.Recordset!BF_COVER
            End If
            If Not glbCompSerial = "S/N - 2380W" Then 'Not for Vitalaire Ticket #11545
                Call updateBenefit(glbLEE_ID, Data1.Recordset!BF_GROUP, TermOrActive, EmployeeBenefitMaster, BCodeCover)
            End If
            'Data1.Refresh
        End If
    End If
End If

Call updBenefitForSalDEPN(glbLEE_ID)

'Vitalaire
If glbCompSerial = "S/N - 2380W" Then Call CalcPP(Trim(clpCode(1).Text), Trim(clpGroup.Text))

DoEvents
Data1.Refresh

'Call Display_Value

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Recordset.Bookmark = bk
End If

Call Display_Value

MsgBox "     Recalculate is Finished.     "

End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalAll_Click()
Dim x%, bk
Dim TermOrActive, BCodeCover
On Error GoTo Recal_Err

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bk = Data1.Recordset.Bookmark
    Data1.Recordset.MoveFirst
End If
Do Until Data1.Recordset.EOF
 
    If glbtermopen Then TermOrActive = "T" Else TermOrActive = "A"
    BCodeCover = Data1.Recordset!BF_BCODE
    If IsNull(Data1.Recordset!BF_COVER) Then
        BCodeCover = BCodeCover & "_"
    Else
        BCodeCover = BCodeCover & "_" & Data1.Recordset!BF_COVER
    End If
    If Not IsNull(Data1.Recordset!BF_GROUP) Then
        If Len(Data1.Recordset!BF_GROUP) > 0 Then
            If Not glbCompSerial = "S/N - 2380W" Then 'Not for Vitalaire Ticket #11545
                Call updateBenefit(glbLEE_ID, Data1.Recordset!BF_GROUP, TermOrActive, EmployeeBenefitMaster, BCodeCover)
            End If
        End If
    End If
Next_Rec:
    Data1.Recordset.MoveNext
Loop

Call updBenefitForSalDEPN(glbLEE_ID)


'vitalaire
If glbCompSerial = "S/N - 2380W" Then Call CalcPP
DoEvents
Data1.Refresh
If glbCElgin Then   'glbCompSerial = "S/N - 2292W" Then
    Call Pause(1)
End If
Data1.Refresh

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Recordset.Bookmark = bk
End If
Call Display_Value
MsgBox "     Recalculate is Finished.     "
Exit Sub

Recal_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdRecalAll", "HRBENFT", "Recalculate")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


'Private Sub Command1_Click()
'Dim rsBenft As New ADODB.Recordset
'Dim SQLQ As String
'
'SQLQ = "SELECT BF_EMPNBR FROM HRBENFT WHERE BF_BCODE = 'DC'"
'rsBenft.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'Do While Not rsBenft.EOF
'    Call updBenefitForSalDEPN(rsBenft("BF_EMPNBR"))
'    rsBenft.MoveNext
'Loop
'rsBenft.Close
'Set rsBenft = Nothing
'End Sub

Private Sub comPreAftTax_Change()
    If comPreAftTax = "Pre Tax" Then
        txtPreAftTax = "P"
    ElseIf comPreAftTax = "After Tax" Then
        txtPreAftTax = "A"
    Else
        txtPreAftTax = ""
    End If
End Sub

Private Sub comPreAftTax_Click()
    If comPreAftTax = "Pre Tax" Then
        txtPreAftTax = "P"
    ElseIf comPreAftTax = "After Tax" Then
        txtPreAftTax = "A"
    Else
        txtPreAftTax = ""
    End If
End Sub

Private Sub comPreAftTax_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comRelation_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comRelation_LostFocus()
Dim tlen As Integer
tlen = Len(comRelation.Text)
If tlen > 10 Then tlen = 10
If tlen >= 1 Then
    lblRel = Left$(comRelation.Text, tlen)
Else
    lblRel = " "
End If

End Sub


Private Sub comRndFactor_Change()
    If Val(txtRoundFactor) <> Val(comRndFactor.ItemData(comRndFactor.ListIndex)) Then
        txtRoundFactor = Val(comRndFactor.ItemData(comRndFactor.ListIndex))
    End If
    Call Set_SalCover
End Sub

Private Sub comRndFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comRndFactor_LostFocus()
    If comRndFactor.ListIndex = -1 Then comRndFactor.ListIndex = 0
    If Val(txtRoundFactor) <> Val(comRndFactor.ItemData(comRndFactor.ListIndex)) Then
        txtRoundFactor = Val(comRndFactor.ItemData(comRndFactor.ListIndex))
    End If
    Call Set_SalCover
End Sub

Private Sub comSalDepn_Change()

frmReleSalaryINFO.Visible = True
If comSalDepn = "Yes" Then
    comRndFactor.Enabled = True
    medMinCover.Enabled = True
    medMaxCover.Enabled = True
    medSalFactor.Enabled = True
    medCovAmount.Enabled = False
    If Not gSec_Inq_Salary Then
        frmReleSalaryINFO.Visible = False
    End If
Else
    comRndFactor.Enabled = False
    medMinCover.Enabled = False
    medMaxCover.Enabled = False
    medSalFactor.Enabled = False
    medCovAmount.Enabled = True
End If
End Sub

'Private Sub comSalDepn_Click()
'comSalDepn_Change
'End Sub

Private Sub comSalDepn_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comSalDepn_Click()
If comSalDepn = "Yes" Then
    lblAP = "P"
    Set_SalCover
Else
End If
txtSalDepn = Left(comSalDepn, 1)
comSalDepn_Change
End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS
Call EMP_Releate
If glbLinamar Then
    SQLQ = "SELECT SUBSTRING(BF_BCODE,4,8) AS BF_SHOWKEY,*"
Else
    SQLQ = "SELECT *"
End If
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HRBENFT "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
Else
    SQLQ = SQLQ & " FROM HRBENFT "
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
End If
Data1.RecordSource = SQLQ
Data1.Refresh
If glbLinamar Then
    SQLQ = "SELECT SUBSTRING(BD_BCODE,4,8) AS BD_SHOWKEY,*"
Else
    SQLQ = "SELECT *"
End If

If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HRBENS "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY BD_BCODE "
Else
    SQLQ = SQLQ & " FROM HRBENS "
    SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY BD_BCODE "
End If
Data3.RecordSource = SQLQ
Data3.Refresh

'Ticket #15898
If FrmBens.Visible = True Then
    Dim xCodeList As String
    Do Until Data1.Recordset.EOF
        xCodeList = xCodeList & "," & Data1.Recordset!BF_BCODE
        Data1.Recordset.MoveNext
    Loop
    If Len(xCodeList) > 0 Then xCodeList = Mid(xCodeList, 2)
    clpBCODE.seleEMPCode = xCodeList 'Pay attention
End If

Set FRS1 = Data1.Recordset.Clone
Set FRS3 = Data3.Recordset.Clone

EERetrieve = True

Screen.MousePointer = DEFAULT


Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HRBENFT", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function

End Function


Private Function Delete_Match_BENS()
Dim EID As Long, BCode As String
Dim SQLQ As String
Dim selDel As New ADODB.Recordset
Delete_Match_BENS = False

On Error GoTo D3Mtch_Err
BCode = Data1.Recordset("BF_BCODE")
If glbtermopen Then
    EID = Data1.Recordset("Term_Seq")
    SQLQ = "DELETE FROM Term_HRBENS "
    SQLQ = SQLQ & " WHERE Term_Seq =" & EID 'Ticket# 7272 missing =
    SQLQ = SQLQ & " AND BD_BCODE= '" & BCode & "'"
    gdbAdoIhr001X.Execute SQLQ
Else
    EID = Data1.Recordset("BF_EMPNBR")
    SQLQ = "DELETE FROM HRBENS "
    SQLQ = SQLQ & " WHERE BD_EMPNBR= " & EID
    SQLQ = SQLQ & " AND BD_BCODE= '" & BCode & "'"
    gdbAdoIhr001.Execute SQLQ
End If

Data3.Refresh

Delete_Match_BENS = True

Exit Function

D3Mtch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete_Match_BENS", "HRBENES", "Select")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Form_Activate()
glbOnTop = "FRMEBENEFITS"

FrmDetails.Top = 2520
FrmDetails.Left = 60

FrmBens.Top = 2880
FrmBens.Left = 0
FrmBens.Width = 10455
FrmBens.Height = 4275 'Ticket #24275 Franks 08/26/2013 '3315

'Franks 03/16/2011
fraCopyBS.Top = 7200 ' 6200 '5640 'Ticket #24275 Franks 08/28/2013
fraCopyBS.Left = 120

Call Form_Resize

If FrmBens.Visible Then
    glbSkip = True
Else
    glbSkip = False
End If

Call LdcomRel 'Ticket #22038 Franks 05/16/2012

fglbNew = False
Call SET_UP_MODE
End Sub

Private Sub Form_Deactivate()
glbSkip = False
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEBENEFITS"

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title, SQLQ '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEBENEFITS"

fUpdable = True

If glbWFC Then
    locWFCEmpID = 0
    glbSkip = False
    Call WFC_ScreenSetup 'Ticket #24275 Franks 08/26/2013
End If

lblTitle(44).Visible = True
dlpDate(2).Visible = True
dlpDate(2).DataField = "BF_CEASEDATE"

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data3.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data3.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS
IfElginLife = False
If glbLinamar Then
    vbxTrueGrid.Columns(0).DataField = "BF_SHOWKEY"
    TblBENS.Columns(0).DataField = "BD_SHOWKEY"
End If
'Call LdcomRel
Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    
    'Ticket #27375 Franks 08/05/2015
    If glbWFC Then
        If glbNoNONE Then
            If glbUNION = "NONE" Then
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
        If glbNoEXEC Then       'Hemu -EXE
            If glbUNION = "EXEC" Then
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
    End If
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    
    'Ticket #27375 Franks 08/05/2015
    If glbWFC Then
        If glbNoNONE Then
            If glbUNIONTe = "NONE" Then
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
        If glbNoEXEC Then
            If glbUNIONTe = "EXEC" Then     'Hemu -EXE
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Unload Me
                Screen.MousePointer = DEFAULT
                Exit Sub
            End If
        End If
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

vbxTrueGrid.FetchRowStyle = True
vbxTrueGrid.MarqueeStyle = 3
TblBENS.FetchRowStyle = True
TblBENS.MarqueeStyle = 3

If glbCElgin Then
    Dim rsTemp As New ADODB.Recordset, rsTemp01 As New ADODB.Recordset
    SQLQ = "SELECT HREMP.ED_EMPNBR, HREMP.ED_PT, HREMP.ED_ORG, HREMP.ED_DOH, HRJOB.JB_GRPCD "
    SQLQ = SQLQ & "FROM (HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) LEFT JOIN HRJOB ON "
    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "WHERE (ED_EMPNBR = " & glbLEE_ID & " And (HR_JOB_HISTORY.JH_CURRENT) <> 0)"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF And rsTemp.BOF Then
        xED_PT = ""
        xED_ORG = ""
        xJB_GRPVD = ""
    Else
        xED_PT = rsTemp("ED_PT")
        xED_ORG = rsTemp("ED_ORG")
        xJB_GRPVD = rsTemp("JB_GRPCD")
    End If
    rsTemp.Close
    Dim xDATE, xYY
    xYY = Year(Now) - 1
    xDATE = CVDate(GetMonth("Dec") & " 1," & Str(xYY))
    SQLQ = "SELECT SH_EMPNBR, SH_SALARY, SH_CURRENT, SH_EDATE, SH_WHRS, SH_SALCD FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & "WHERE (SH_EMPNBR = " & glbLEE_ID & " And SH_EDATE <= " & Date_SQL(xDATE) & ") "
    SQLQ = SQLQ & "ORDER BY SH_EDATE DESC "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xCovAmt = 0
    If Not (rsTemp.EOF And rsTemp.BOF) Then
        xCovAmt = rsTemp("SH_SALARY")
        If rsTemp("SH_SALCD") = "H" Then
            If rsTemp("SH_WHRS") > 0 Then
                xCovAmt = xCovAmt * rsTemp("SH_WHRS") * 52
            End If
        ElseIf rsTemp("SH_SALCD") = "M" Then
            xCovAmt = xCovAmt * 12
        ElseIf rsTemp("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                xCovAmt = xCovAmt * 366
            Else
                xCovAmt = xCovAmt * 365
            End If
        End If
    End If
End If
'If glbVadim Then
'    lblTitle(45).Visible = True
'    txtPayrollID.Visible = True
'End If

If glbVadim Then
    lblTitle(46).Visible = True
    medRateLevel.Visible = True
End If

If glbCompSerial = "S/N - 2290W" Then 'Ticket #22464 - Goodmans
    lblTitle(46).Visible = True
    medRateLevel.Visible = True
    lblTitle(46).Caption = "Sequence #"
    medRateLevel.Tag = "10-Sequence #"
    medRateLevel.MaxLength = 2
End If

TblBENS.Visible = False
FrmBens.Visible = False

'Add by Franks Dec 6,02 #3298 , allow the % Company and % Employee to exceed 100%
medPPE.Enabled = True

Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) < 1 Then Exit Sub

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Benefits - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

If glbLinamar Then
    clpCode(1).MaxLength = 8
    clpBCODE.MaxLength = 8
    txtDiv = Left(lblEENum, 3)
    txtEmployeeID.Visible = True
    txtSortCode.Visible = True
    lblTitle(34).Visible = True
    lblTitle(35).Visible = True
    lblTitle(36).Visible = True
    lblTitle(36).Caption = "Benefit Code/Class"
    lblDOH.Visible = True
    lblYear.Visible = True
    txtSortCode.DataField = "BF_SORTCODE"
    txtSortCode.Tag = "00-Benefit Code/Class"
    txtEmployeeID.DataField = "BF_EMPLOYEEID"
    lblTitle(40).Visible = False  'Not available for Linamar
    clpGroup.Visible = False 'Not available for Linamar
End If


Call lblAP_Change
'---------------------------- Read BENS

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Call Display_Value

FrmDetails.Enabled = False

Call INI_Controls(Me)

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call SET_UP_MODE

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
glbSkip = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)


End Sub

Private Sub Form_Resize()
If FrmDetails.Visible = False Then scrControl.Visible = False: Exit Sub

If Me.Height >= vbxTrueGrid.Height + FrmDetails.Height + panControls.Height + panEEDESC.Height + 230 Then
    scrControl.Value = 0
    FrmDetails.Top = vbxTrueGrid.Height + panEEDESC.Height - 200
    scrControl.Visible = False
    Exit Sub
End If
If Me.Height < vbxTrueGrid.Height + scrControl.Top + panControls.Height Then Exit Sub
scrControl.Visible = True
scrControl.Max = vbxTrueGrid.Height + FrmDetails.Height + panControls.Height - Me.Height + 250
scrControl.Left = Me.Width - scrControl.Width - 120
scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 500 '- 400
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEBENEFITS = Nothing  'carmen may 2000
Call NextForm
End Sub



Private Sub lblAP_Change()
If lblAP = "A" Or lblAP = "" Then
    optActual(0).Value = True
Else
    optActual(1).Value = True
End If
If optActual(1).Value = True Then
    txtPer.Enabled = True
    medUnitCost.Enabled = True
    medTCost.Enabled = False
Else
    txtPer.Enabled = False
    medUnitCost.Enabled = False
    medTCost.Enabled = True
End If
End Sub

Private Sub lblRel_Change()
comRelation.Text = lblRel
End Sub

Private Sub LdcomRel()
'comRelation.Clear
'comRelation.AddItem "Wife"
'comRelation.AddItem "Husband"
'comRelation.AddItem "Common Law"
'comRelation.AddItem "Brother"
'comRelation.AddItem "Sister"
'comRelation.AddItem "Daughter"
'comRelation.AddItem "Son"
'comRelation.AddItem "Father"
'comRelation.AddItem "Mother"
'comRelation.AddItem "Fiancee"
'comRelation.AddItem "Fiance"
'comRelation.AddItem "Spouse"
'comRelation.AddItem "Estate"
'comRelation.AddItem "Other"
'comRelation.AddItem "Uncle"
'comRelation.AddItem "Aunt"
'comRelation.AddItem "Couple"
'comRelation.AddItem "Children"
'comRelation.AddItem "Parents"
'comRelation.AddItem "Ex-Spouse"

'WFC Pension Outstanding Tasks By Dec2109.doc in W:\2008 Projects\Pension\Pension Phase II
'Beneficiary screen: Relationship drop down should be alphabetical but Spouse should default (all users)
If glbWFC And glbEmpCountry = "CANADA" Then
    ''Ticket #22009 Franks 05/11/2012
    'comRelation.AddItem "Child"
    'comRelation.AddItem "Spouse"
    'Ticket #22038 Franks 05/15/2012
    comRelation.Clear
    comRelation.AddItem "Aunt"
    comRelation.AddItem "Brother"
    comRelation.AddItem "Children"
    comRelation.AddItem "Common Law"
    'comRelation.AddItem "Couple"
    comRelation.AddItem "Daughter"
    comRelation.AddItem "Estate"
    comRelation.AddItem "Ex-Spouse"
    comRelation.AddItem "Father"
    comRelation.AddItem "Fiance"
    comRelation.AddItem "Fiancee"
    'comRelation.AddItem "Husband"
    comRelation.AddItem "Mother"
    comRelation.AddItem "Other"
    comRelation.AddItem "Parents"
    comRelation.AddItem "Sister"
    comRelation.AddItem "Son"
    comRelation.AddItem "Spouse"
    comRelation.AddItem "Uncle"
    'comRelation.AddItem "Wife"
Else
    comRelation.Clear
    comRelation.AddItem "Aunt"
    comRelation.AddItem "Brother"
    comRelation.AddItem "Children"
    comRelation.AddItem "Common Law"
    comRelation.AddItem "Couple"
    comRelation.AddItem "Daughter"
    comRelation.AddItem "Estate"
    comRelation.AddItem "Ex-Spouse"
    comRelation.AddItem "Father"
    comRelation.AddItem "Fiance"
    comRelation.AddItem "Fiancee"
    comRelation.AddItem "Husband"
    comRelation.AddItem "Mother"
    comRelation.AddItem "Other"
    comRelation.AddItem "Parents"
    comRelation.AddItem "Sister"
    comRelation.AddItem "Son"
    comRelation.AddItem "Spouse"
    comRelation.AddItem "Uncle"
    comRelation.AddItem "Wife"
End If

End Sub

Private Sub lblRound_Change()
If lblRound = "R" Then
    optRound(0) = True
Else
    optRound(1) = True
End If
End Sub

Private Sub medCompCost_LostFocus()
If Len(medCompCost) > 0 Then
    If IsNumeric(medCompCost) Then
        If Len(medTCost) > 0 Then
            If IsNumeric(medTCost) Then
                medPPComp = medCompCost / medTCost
            End If
        End If
    End If
End If

End Sub

Private Sub medCovAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medCovAmount_LostFocus()
If Len(medCovAmount) > 0 Then
    If Not IsNumeric(medCovAmount) Then
        MsgBox "You must enter dollar value"
        medCovAmount.SetFocus
        Exit Sub
    End If
End If
'Call setTotal
Call Set_SalCover
End Sub


Private Sub medEECost_LostFocus()
If Len(medEECost) > 0 Then
    If IsNumeric(medEECost) Then
        If Len(medTCost) > 0 Then
            If IsNumeric(medTCost) Then
                If medTCost = 0 Then
                MsgBox "Can't divide by 0"
                Else
                    medPPE = medEECost / medTCost
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub medMaxAmnt_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medMaxAmnt_LostFocus()
If Len(medMaxAmnt) > 0 Then
    If Not IsNumeric(medMaxAmnt) Then
        MsgBox "You must enter dollar value"
        medMaxAmnt.SetFocus
        Exit Sub
    End If
End If

End Sub


Private Sub medMaxCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMaxCover_LostFocus()
If Len(Trim(medMaxCover)) = 0 Then medMaxCover = 0 'Jaddy 11/3/99
Call Set_SalCover
End Sub

Private Sub medMCCOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMCCOST_LostFocus()
If Len(medMCCOST) = 0 Or medMCCOST = "." Then medMCCOST = 0
'Frank Nov 28, 01 for Casey House
If glbCompSerial = "S/N - 2214W" And optActual(0) Then
    Call CaseyHouseBenefit
End If
'Frank Nov 28, 01 for Casey House
End Sub

Private Sub medMECOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMECOST_LostFocus()
If Len(medMECOST) = 0 Or medMECOST = "." Then medMECOST = 0
'Frank Nov 28, 01 for Casey House
If glbCompSerial = "S/N - 2214W" And optActual(0) Then
    Call CaseyHouseBenefit
End If
'Frank Nov 28, 01 for Casey House
End Sub

Private Sub Set_SalCover()
Dim xSalFactor, xRndFactor, xMaxCover, xMinCover, xCovAmount
Dim xSalary
Dim CostINFO As BenefitCost
Dim xDecimal
    If comSalDepn = "Yes" Then
        If glbtermopen Then
            xSalary = CrtSalary(glbTERM_Seq)
        Else
            'If Not IfElginLife Then
                xSalary = CrtSalary(glbLEE_ID)
            'Else
            '    xSalary = xCovAmt
            'End If
        End If
        CostINFO = CrtBeneCost(glbLEE_ID, xSalary, clpGroup.Text, clpCode(1).Text)
        xSalary = CostINFO.Salary
        
        xSalFactor = Val(medSalFactor)
        xRndFactor = Val(txtRoundFactor)
        xMaxCover = Val(medMaxCover)
        xMinCover = Val(medMinCover)
        xCovAmount = xSalary * xSalFactor
        'rounding moved before setting to min or max by Bryan 18/Oct/05 Ticket#9487
        If xRndFactor = 0 Then xRndFactor = 0.01
        
        'If Rounding factor is 1 and Next, and Coverage Amount is a whole # then do not add 0.5
        If xRndFactor = 1 And optRound(0) = False Then
            xDecimal = xCovAmount - Int(xCovAmount)
            If xDecimal = 0 Then
                xCovAmount = Round(xCovAmount / xRndFactor) * xRndFactor
            Else
                xCovAmount = Round(xCovAmount / xRndFactor + IIf(optRound(0) = True, 0, 0.5)) * xRndFactor
            End If
        Else
            'Ticket #18465 - If NEXT and evenly divisible (whole #) then do not round to NEXT
            If optRound(0) = False Then
                xDecimal = (xCovAmount / xRndFactor) - Int(xCovAmount / xRndFactor)
                If xDecimal = 0 Then
                    xCovAmount = Round(xCovAmount / xRndFactor) * xRndFactor
                Else
                    xCovAmount = Round(xCovAmount / xRndFactor + IIf(optRound(0) = True, 0, 0.5)) * xRndFactor
                End If
            Else
                xCovAmount = Round(xCovAmount / xRndFactor + IIf(optRound(0) = True, 0, 0.5)) * xRndFactor
            End If
        End If
        
        If xMinCover <> 0 And xCovAmount < xMinCover Then xCovAmount = xMinCover
        If xMaxCover <> 0 And xCovAmount > xMaxCover Then xCovAmount = xMaxCover
        
        medCovAmount = xCovAmount
        If CostINFO.Type = "M" Or CostINFO.Type = "W" Then  'Ticket #25235 - For weekly too * 12 even though the Covrg Amt is Weekly
            AnCoverAmt = xCovAmount * 12
        'Ticket #25235 - This is not working so Jerry and I decided to use the above * 12 which gives the right result for the client
        'ElseIf CostINFO.Type = "W" Then     'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
        '    AnCoverAmt = xCovAmount * 52
        Else
            AnCoverAmt = xCovAmount
        End If
        
    End If

    Call setTotal

End Sub

Private Sub medMinCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMinCover_LostFocus()
If Len(Trim(medMinCover)) = 0 Then medMinCover = 0 'Jaddy 11/3/99
Call Set_SalCover
End Sub

Private Sub medPayPeriodAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPayPeriodAmount_LostFocus()
If Len(Trim(medPayPeriodAmount)) = 0 And medPayPeriodAmount.DataChanged Then medPayPeriodAmount = 0  'Jaddy 11/3/99
'Call setTotal
Call Set_SalCover
If glbWFC Then 'Ticket #22964 Franks 12/17/2012
    Call WFCCalUnitCost4DCPP
End If
End Sub

Private Sub WFCCalUnitCost4DCPP() 'Ticket #22964 Franks 12/17/2012
    If clpCode(1).Text = "DCPP" Then
        If IsNumeric(medPayPeriodAmount.Text) Then
            medUnitCost.Text = medPayPeriodAmount / 100
        End If
    End If
End Sub

Private Sub medPPComp_GotFocus()
Call SetPanHelp(ActiveControl)
medPPComp = Val(medPPComp) * 100
End Sub

Private Sub medPPComp_KeyUp(KeyCode As Integer, Shift As Integer)
'Comment by Franks Dec 6,02 #3298 , allow the % Company and % Employee to exceed 100%
'If Val(medPPComp) > 100 Then medPPComp = 100
End Sub


Private Sub medPPComp_LostFocus()
    medPPComp = Val(medPPComp) / 100
    'Comment by Franks Dec 6,02 #3298 , allow the % Company and % Employee to exceed 100%
    'medPPE = 1 - Val(medPPComp)
    'Call setTotal
    Call Set_SalCover
End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
medPPE = Val(medPPE) * 100
End Sub


Private Sub medPPE_LostFocus()
If Len(medPPE) > 0 Then
    If IsNumeric(medPPE) Then
        medEECost = Val(medTCost) * Val(medPPE) / 100
        medPPE = Val(medPPE) / 100
        'Add by Franks Dec 6,02 #3298 , allow the % Company and % Employee to exceed 100%
        'Call setTotal
        Call Set_SalCover
        
        medEECost.Visible = True
    End If
End If
End Sub

Private Sub medRateLevel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSalFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSalFactor_LostFocus()
Call Set_SalCover

End Sub

Private Sub MedSplitPc_GotFocus()
If MedSplitPc = "" Then
   MedSplitPc = 0
End If

MedSplitPc = MedSplitPc * 100
Call SetPanHelp(ActiveControl)

End Sub


Private Sub MedSplitPc_LostFocus()
If Not IsNumeric(MedSplitPc) Then
   MedSplitPc = 0
End If
MedSplitPc = MedSplitPc / 100

End Sub



Private Sub medTCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medTCost_LostFocus()
'Call setTotal
Call Set_SalCover
End Sub

Private Sub medUnitCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUnitCost_LostFocus()
'Call setTotal
Call Set_SalCover
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optActual_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optActual_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If optActual(1).Value = True Then
    txtPer.Enabled = True
    medUnitCost.Enabled = True
    medTCost.Enabled = False
Else
    If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then txtPer = 0
    txtPer.Enabled = False
    medUnitCost.Enabled = False
    medTCost.Enabled = True
End If
End Sub

Private Sub optActual_LostFocus(Index As Integer)
If optActual(0).Value = True Then lblAP = "A"
If optActual(1).Value = True Then lblAP = "P"
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

fUPMode = TF    ' update mode
FrmDetails.Enabled = TF
FrmBens.Enabled = TF
If fEBGroup = "NOGROUP" Then
    clpGroup.Enabled = False
Else
    clpGroup.Enabled = True
End If
If FrmDetails.Visible = True Then
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        cmdRecalAll.Enabled = False
        cmdRecal.Enabled = False
        cmdBens.Enabled = False
    Else
        cmdRecalAll.Enabled = True
        cmdRecal.Enabled = True
        
        If gSec_Inq_Beneficiary Then
            cmdBens.Enabled = True
        Else
            cmdBens.Enabled = False
        End If
    End If
    
    If glbtermopen Then cmdRecalAll.Visible = False
End If
If Not gSec_Inq_Salary Then
    comSalDepn.Enabled = False
End If
If Not gSec_Upd_Benefits Then
    cmdRecalAll.Enabled = False
    cmdRecal.Enabled = False
    
    If Not gSec_Inq_Beneficiary Then
        cmdBens.Enabled = False
    End If
End If
If glbCElgin Then cmbDWM.Enabled = False
End Sub

Private Sub optActual_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If optActual(1).Value = True Then
    txtPer.Enabled = True
    medUnitCost.Enabled = True
    medTCost.Enabled = False
Else
    If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then txtPer = 0
    txtPer.Enabled = False
    medUnitCost.Enabled = False
    medTCost.Enabled = True
End If
End Sub

Private Sub optRound_LostFocus(Index As Integer)
Call Set_SalCover
lblRound = IIf(optRound(0), "R", "N")
End Sub

Private Sub optRound_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Set_SalCover
End Sub

Private Sub scrControl_Change()
FrmDetails.Top = 250 + vbxTrueGrid.Height - scrControl.Value * 2.5
End Sub

Private Sub TblBENS_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub TblBENS_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo Eh
    'added by Bryan 18/Jan/06 Ticket#10222
    FRS3.Requery
    'If Not FRS3.EOF Then
    '    FRS3.Bookmark = Bookmark
    'End If
    'change row colour
'    If FRS("BD_FREEZE") = True Then
'        RowStyle.ForeColor = vbRed
'    End If
    
Eh:
    Exit Sub
End Sub

Private Sub TblBENS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub TblBENS_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
        
        If TblBENS.Tag = "ASC" Then
            TblBENS.Tag = "DESC"
        Else
            TblBENS.Tag = "ASC"
        End If
        
        If glbLinamar Then
            SQLQ = "SELECT SUBSTRING(BD_BCODE,4,8) AS BD_SHOWKEY,*"
        Else
            SQLQ = "SELECT *"
        End If
        
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_HRBENS "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = SQLQ & " FROM HRBENS "
            SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY  " & TblBENS.Columns(ColIndex).DataField & " " & TblBENS.Tag
        
    
        Data3.RecordSource = SQLQ
        Data3.Refresh
        Set FRS3 = Data3.Recordset.Clone
        TblBENS.FetchRowStyle = True

End Sub

Private Sub TblBENS_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
Call getCodes

Call WFCPenFieldsEnable(False) 'Ticket #24275 Franks 08/27/2013

End Sub

Private Sub txtBeneName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Hemu - 10/03/2003 Begin - Ticket # 4728, This was commented in 6.0 and 7.0
'Private Sub CalElgin()
'    Flag1 = False
'    Flag2 = False
'    Flag3 = False
'    If UCase(clpCode(1).Text) = "LIFE" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
'        Flag1 = True
'    End If
'    ' dkostka - 06/11/2002 - Add new criteria, discussed w/ Paul and Dorothy in conference call (#2367)
'    If UCase(clpCode(1).Text) = "LIFE" And xED_PT = "FT" And xED_ORG = "1" Then
'        Flag1 = True
'    End If
'    ' dkostka - 06/11/2002 - end change
'    If UCase(clpCode(1).Text) = "AD&D" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
'        Flag2 = True
'    End If
'    If UCase(clpCode(1).Text) = "AD&D" And xED_PT = "FT" And xED_ORG = "3" And (xJB_GRPVD = "ONA") Then
'        Flag3 = True
'    End If
'    If Flag1 Or Flag2 Or Flag3 Then
'        IfElginLife = True
'        txtWaitPeriod = 90
'        Call txtWaitPeriod_LostFocus
'
'        txtSalDepn = "Y"
'        comSalDepn.Enabled = False
'
'        txtCovType = "Y"
'        medMinCover = 0
'        ' dkostka - 06/11/2002 - Only fill in max coverage if they left it blank.
'        If Val(medMaxCover) = 0 Then medMaxCover = 200000
'        ' dkostka - 06/11/2002 - end change
'
'        medSalFactor = 2
'        lblRound = "N"
'        comRndFactor = 1000
'        If Flag1 Then
'            medUnitCost = 0.372
'        End If
'        If Flag2 Or Flag3 Then
'            medUnitCost = 0.025
'        End If
'        txtPer = 1000
'        medPPComp = 1
'        medPPE = 0
'
'        lblAP = "P"
'        Call Set_SalCover
'
'        ' dkostka - 02/18/2002 - Changed next line to only force Y for LIFE, not AD&D.
'        '   Paul requested this.
'        If UCase(Data1.Recordset("BF_BCODE")) <> "AD&D" Then txtTAXBEN = "Y"
'
'    Else
'        IfElginLife = False
'        ' dkostka - 02/12/2002 - Added code to re-enable salary dependent y/n box.  If we don't
'        '   do this it never gets enabled and they can't change any records after looking at
'        '   an AD&D record once.
'        comSalDepn.Enabled = True
'    End If
'
'End Sub
'Hemu - 10/02/2003 End - Ticket # 4728

Private Sub CalElgin()
    Flag1 = False
    Flag2 = False
    Flag3 = False
    If UCase(clpCode(1).Text) = "LIFE" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
        Flag1 = True
    End If
    ' dkostka - 06/11/2002 - Add new criteria, discussed w/ Paul and Dorothy in conference call (#2367)
    If UCase(clpCode(1).Text) = "LIFE" And xED_PT = "FT" And xED_ORG = "1" Then
        Flag1 = True
    End If
    ' dkostka - 06/11/2002 - end change
    If UCase(clpCode(1).Text) = "AD&D" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
        Flag2 = True
    End If
    If UCase(clpCode(1).Text) = "AD&D" And xED_PT = "FT" And xED_ORG = "3" And (xJB_GRPVD = "ONA") Then
        Flag3 = True
    End If
    If Flag1 Or Flag2 Or Flag3 Then
        IfElginLife = True
        txtSalDepn = "Y"
        comSalDepn.Enabled = False
        lblAP = "P"
        Call Set_SalCover
    Else
        IfElginLife = False
        ' dkostka - 02/12/2002 - Added code to re-enable salary dependent y/n box.  If we don't
        '   do this it never gets enabled and they can't change any records after looking at
        '   an AD&D record once.
        comSalDepn.Enabled = True
    End If

End Sub

'Hemu - 10/03/2003 Begin - Ticket # 4728, This was commented in 6.0 and 7.0
'Private Sub CalElginAll()
'
'    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
'        Data1.Recordset.MoveFirst
'    Else
'        Exit Sub
'    End If
'
'    TblBENS.MoveFirst
'
'    Do While Not Data1.Recordset.EOF 'BF_BCODE
'
'        Flag1 = False
'        Flag2 = False
'        Flag3 = False
'        If UCase(Data1.Recordset("BF_BCODE")) = "LIFE" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
'            Flag1 = True
'        End If
'        ' dkostka - 06/11/2002 - Add new criteria, discussed w/ Paul and Dorothy in conference call (#2367)
'        If UCase(Data1.Recordset("BF_BCODE")) = "LIFE" And xED_PT = "FT" And xED_ORG = "1" Then
'            Flag1 = True
'        End If
'        ' dkostka - 06/11/2002 - end change
'        If UCase(Data1.Recordset("BF_BCODE")) = "AD&D" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
'            Flag2 = True
'        End If
'        If UCase(Data1.Recordset("BF_BCODE")) = "AD&D" And xED_PT = "FT" And xED_ORG = "3" And (xJB_GRPVD = "ONA") Then
'            Flag3 = True
'        End If
'        If Flag1 Or Flag2 Or Flag3 Then
'            IfElginLife = True
'            txtWaitPeriod = 90
'            Call txtWaitPeriod_LostFocus
'
'            txtSalDepn = "Y"
'            comSalDepn.Enabled = False
'
'            txtCovType = "Y"
'            medMinCover = 0
'            ' dkostka - 06/11/2002 - Only fill in max coverage if they left it blank.
'            If Val(medMaxCover) = 0 Then medMaxCover = 200000
'            ' dkostka - 06/11/2002 - end change
'
'            medSalFactor = 2
'            lblRound = "N"
'            comRndFactor = 1000
'            medUnitCost = 0.372
'            txtPer = 1000
'            medPPComp = 1
'            medPPE = 0
'
'            lblAP = "P"
'            Call Set_SalCover
'
'            ' dkostka - 02/18/2002 - Changed next line to only force Y for LIFE, not AD&D.
'            '   Paul requested this.
'            If UCase(Data1.Recordset("BF_BCODE")) <> "AD&D" Then txtTAXBEN = "Y"
'
'        Else
'            IfElginLife = False
'        End If
'        Data1.Recordset.MoveNext
'    Loop
'End Sub
'Hemu - 10/02/2003 End - Ticket # 4728

Private Sub txtCovType_GotFocus()
Call SetPanHelp(ActiveControl)
OldCovType = txtCovType
End Sub

Private Sub txtCovType_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCovType_LostFocus()
If OldCovType <> txtCovType Then
    Call ResetValues
End If
End Sub

Private Sub txtSCert_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSComp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSortCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDWM_Change()
cmbDWM.ListIndex = -1
Select Case txtDWM
Case "D"
    cmbDWM.ListIndex = 0
Case "W"
    cmbDWM.ListIndex = 1
Case "M"
    cmbDWM.ListIndex = 2
End Select
End Sub
Private Sub txtPerorDoll_Change()
cmbPerOrDoll.ListIndex = -1
Select Case txtPerOrDoll
Case "D"
    cmbPerOrDoll.ListIndex = 0
Case "P"
    cmbPerOrDoll.ListIndex = 1
End Select
End Sub
Private Sub txtEmployeeID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPer_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub




Private Sub txtPer_LostFocus()
If Len(Trim(txtPer)) = 0 Then txtPer = 0 'Jaddy 11/3/99
'Call setTotal
Call Set_SalCover
End Sub

Private Sub txtPolicy_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPreAftTax_Change()

If txtPreAftTax = "P" Then
    comPreAftTax.ListIndex = 0
ElseIf txtPreAftTax = "A" Then
    comPreAftTax.ListIndex = 1
Else
    comPreAftTax = ""
End If

End Sub

Private Sub txtRoundFactor_Change()
    Dim c As Long
    
        For c = 0 To comRndFactor.ListCount - 1
            If comRndFactor.ItemData(c) = Val(txtRoundFactor) Then
                comRndFactor.ListIndex = c
                Exit For
            End If
        Next c

End Sub

Private Sub txtSalDepn_Change()
comSalDepn = IIf(txtSalDepn = "Y", "Yes", "No")
End Sub

Private Sub txtSPlan_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTAXBEN_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTAXBEN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtWaitPeriod_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtWaitPeriod_LostFocus()
Dim rsEMP As New ADODB.Recordset
Dim xPER
Dim xDATE, xYY, xMM, xDD
Dim xDateAge65

If IsNumeric(txtWaitPeriod) Then
    rsEMP.Open "SELECT ED_DOH, ED_USRDAT1, ED_DOB FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        If IsDate(rsEMP("ED_DOH")) Then
            If Not glbCElgin Then
                 xPER = "M"
                 If cmbDWM <> "" Then
                    xPER = Left(cmbDWM, 1)
                    If xPER = "W" Then xPER = "ww"
                 End If
                 If Actn = "A" Then 'Jerry said the Effective Date should be recalculated only when a record is Added
                    'Ticket #24203 - Family Day Care Services
                    'Ticket #21504 - Kerry's Place
                    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2436W" Then
                        If IsDate(rsEMP("ED_USRDAT1")) Then
                            dlpDate(0) = DateAdd(xPER, Val(txtWaitPeriod), rsEMP("ED_USRDAT1"))
                        Else
                            MsgBox lStr("User Defined date") & " is missing on the employee's Status/Dates screen to compute Effective Date, instead using " & lStr("Original Hire Date") & ".", vbInformation
                            dlpDate(0) = DateAdd(xPER, Val(txtWaitPeriod), rsEMP("ED_DOH"))
                        End If
                    Else
                        'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
                        If glbCompSerial = "S/N - 2420W" And clpCode(1).Text = "PEN" Then
                            If Day(rsEMP("ED_DOH")) = 1 Then
                                dlpDate(0) = DateAdd(xPER, Val(txtWaitPeriod), rsEMP("ED_DOH"))
                            Else
                                xDATE = MonthLastDate(rsEMP("ED_DOH"))
                                dlpDate(0) = DateAdd("d", 1, CVDate(xDATE))
                            End If
                        Else
                            dlpDate(0) = DateAdd(xPER, Val(txtWaitPeriod), rsEMP("ED_DOH"))
                        End If
                    End If
                 End If
                 If glbLinamar Then
                    If cmbDWM.ListIndex >= 0 And cmbDWM.ListIndex <= 2 Then
                        dlpDate(0) = DateAdd(IIf(cmbDWM.ListIndex = 1, "ww", Left(cmbDWM, 1)), Val(txtWaitPeriod), rsEMP("ED_DOH"))
                    End If
                 End If
            Else
                'If Actn = "A" Then
                    xDATE = DateAdd("d", Val(txtWaitPeriod), rsEMP("ED_DOH"))
                    xDD = Day(CVDate(xDATE))
                    If xDD > 15 Then
                        xDATE = DateAdd("d", -(xDD - 1), CVDate(xDATE))
                        dlpDate(0).Text = DateAdd("m", 1, CVDate(xDATE))
                    Else
                        dlpDate(0).Text = CVDate(xDATE)
                    End If
                'End If
            End If
        End If
    End If
End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo Eh
    'added by Bryan 18/Jan/06 Ticket#10222
    FRS1.Requery
    'If Not FRS1.EOF Then
    '    FRS1.Bookmark = Bookmark
    'End If
    'change row colour
'    If FRS("BD_FREEZE") = True Then
'        RowStyle.ForeColor = vbRed
'    End If
    
Eh:
    Exit Sub
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
    
    If glbLinamar Then
        SQLQ = "SELECT SUBSTRING(BF_BCODE,4,8) AS BF_SHOWKEY,*"
    Else
        SQLQ = "SELECT *"
    End If
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_HRBENFT "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = SQLQ & " FROM HRBENFT "
        SQLQ = SQLQ & " WHERE BF_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
    Set FRS1 = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

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

Private Sub setTotal()
Dim NomalCCost
Dim AMT, tmpTotalCost As Double

'On Error GoTo Eh
If IsNumeric(medTCost) Then
    tmpTotalCost = medTCost 'Ticket #13776
End If

If comSalDepn = "Yes" Then
    AMT = AnCoverAmt
Else
    AMT = medCovAmount
End If

If optActual(1) = True Then
    'Hemu - 09/11/2003 - Ticket 4728, Commented the Elgin code like 6.0 and 7.0
    'If Not glbCElgin Then
    If glbCompSerial = "S/N - 2439W" And clpCode(1).Text = "STD" Then   'OK Tire  - Ticket #22562 Franks 09/26/2012, Ticket #22660 - for STD only
        If Val(txtPer) > 0 Then tmpTotalCost = (Val(AMT) / 52 / Val(txtPer)) * Val(medUnitCost) Else tmpTotalCost = 0
    Else
        If Val(txtPer) > 0 Then tmpTotalCost = (Val(AMT) / Val(txtPer)) * Val(medUnitCost) Else tmpTotalCost = 0
    End If
    'Else
    '    If IfElginLife Then
    '        If (Flag2 Or Flag3) Then
    '            If Val(txtPer) > 0 Then tmpTotalCost = ((Val(AMT) / Val(txtPer)) * Val(medUnitCost)) * 12 * 1.08
    '        Else
    '            If Val(txtPer) > 0 Then tmpTotalCost = (Val(AMT) / Val(txtPer)) * Val(medUnitCost) * 12
    '        End If
    '    Else
    '        If Val(txtPer) > 0 Then tmpTotalCost = (Val(AMT) / Val(txtPer)) * Val(medUnitCost) Else tmpTotalCost = 0
    '    End If
    'End If
End If

If glbCompSerial = "S/N - 2262W" Then 'Wellington - Ticket #10718
    If clpCode(1).Text = "5ADB" Or clpCode(1).Text = "5GRB" Or clpCode(1).Text = "5LTB" Or _
        clpCode(1).Text = "6ADB" Or clpCode(1).Text = "6GRB" Or clpCode(1).Text = "6LTB" Or _
        clpCode(1).Text = "8GRB" Or clpCode(1).Text = "4ADW" Or clpCode(1).Text = "4GRW" Or _
        clpCode(1).Text = "4LTW" Or clpCode(1).Text = "1GRB" Then
            
        tmpTotalCost = Round(Val(tmpTotalCost), 2)
    End If
End If

'Ticket #20872 Franks 09/28/2011
If clpCode(1).Text = "OMER" And OMER_UseCostTable Then
    tmpTotalCost = EmpOmersCalculate(glbLEE_ID, "OMER", "N")
    medTCost = tmpTotalCost
End If

medEECost = Val(tmpTotalCost) * Val(medPPE)
medCompCost = Val(tmpTotalCost) * Val(medPPComp)

If glbCompSerial = "S/N - 2262W" Then 'Wellington
    If Right(Round(CStr(Val(medEECost) / 12), 3), 1) = 5 Then
        If clpCode(1).Text = "5ADB" Or clpCode(1).Text = "5GRB" Or clpCode(1).Text = "5LTB" Or _
            clpCode(1).Text = "6ADB" Or clpCode(1).Text = "6GRB" Or clpCode(1).Text = "6LTB" Or _
            clpCode(1).Text = "8GRB" Or clpCode(1).Text = "4ADW" Or clpCode(1).Text = "4GRW" Or _
            clpCode(1).Text = "4LTW" Or clpCode(1).Text = "1GRB" Or clpCode(1).Text = "GRLB" Then
            
            medMECOST = Val(medEECost) / 12
        Else
            medMECOST = Round((Val(medEECost) / 12) - 0.005, 2)
        End If
    Else
        medMECOST = Val(medEECost) / 12
    End If
    If Right(Round(CStr(Val(medCompCost) / 12), 3), 1) = 5 Then
        If clpCode(1).Text = "5ADB" Or clpCode(1).Text = "5GRB" Or clpCode(1).Text = "5LTB" Or _
            clpCode(1).Text = "6ADB" Or clpCode(1).Text = "6GRB" Or clpCode(1).Text = "6LTB" Or _
            clpCode(1).Text = "8GRB" Or clpCode(1).Text = "4ADW" Or clpCode(1).Text = "4GRW" Or _
            clpCode(1).Text = "4LTW" Or clpCode(1).Text = "1GRB" Or clpCode(1).Text = "GRLB" Then
            
            medMCCOST = Val(medCompCost) / 12
        Else
            medMCCOST = Round((Val(medCompCost) / 12) - 0.005, 2)
        End If
    Else
        medMCCOST = Val(medCompCost) / 12
    End If

' danielk - 12/30/2002 - Priority C changes for 7.0
' Franks Sep 4,2002 #2774 for Mitchell Plastic
ElseIf (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then
    'medMECOST = Round(Val(medEECost) / 12, 2)
    'medMCCOST = Round(Val(medCompCost) / 12, 2)
    medMECOST = Val(medEECost) / 12
    medMCCOST = Val(medCompCost) / 12
    
    'Hemu - Begin - Surrey Place -------------------
    'Jaddy Changed by Jerry request
    If glbCompSerial = "S/N - 2347W" And clpCode(1).Text = "VH" Then
        Dim FTE
        FTE = GetJHData(glbLEE_ID, "JH_FTENUM", 1)
        NomalCCost = Val(tmpTotalCost) * Val(medPPComp)
        medCompCost = FTE * NomalCCost
        medEECost = Val(tmpTotalCost) * Val(medPPE) + (1 - FTE) * NomalCCost
        medMCCOST = Val(medCompCost) / 12
        medMECOST = Val(medEECost) / 12
    End If
    'Hemu - End -------------------------------------
Else
    medMECOST = 0
    medMCCOST = 0
End If

'Added by Bryan Aug 21, 2007 - Ticket#13546
If comSalDepn.Text = "Yes" Then
    If Val(medPPComp.Text) + Val(medPPE.Text) = 1 Then 'If the total is 100%
        medTCost.Text = tmpTotalCost
    Else 'if the total is not  100%
        medTCost.Text = Val(medCompCost.Text) + Val(medEECost.Text)
    End If
Else
    medTCost.Text = tmpTotalCost
End If

'Vitalaire
If glbCompSerial = "S/N - 2380W" Then
    Call CalcPP(Trim(clpCode(1).Text), Trim(clpGroup.Text))
    Select Case Trim(clpGroup.Text)
    Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"   'Ticket #18963, Ticket #24537 - more codes
        medPayPeriodAmount = Val(medEECost.Text) / 52
    Case Else
        If (medPayPeriodAmount = "0" Or medPayPeriodAmount = "" Or IsNull(medPayPeriodAmount)) Then
            medPayPeriodAmount = Val(medMECOST) / 2
        End If
    End Select
End If

'Franks Sep 4,2002 #2774 for Mitchell Plastic
End Sub


Private Sub UpdCodes1()
If glbLinamar Then
    rsDATA!BF_BCODE = clpCode(1).TransDiv & clpCode(1).Text
Else
    rsDATA!BF_BCODE = clpCode(1).Text
End If
End Sub
Private Sub UpdCodes2()
If glbLinamar Then
    rsDATA3!BD_BCODE = clpCode(1).TransDiv & clpBCODE.Text
Else
    rsDATA3!BD_BCODE = clpBCODE.Text
End If
End Sub
Sub getCodes()
If glbLinamar Then
    clpCode(1).TransDiv = Right(lblEEID, 3)
    clpBCODE.TransDiv = Right(lblEEID, 3)
End If
If Data1.Recordset.EOF Then
    clpCode(1).Text = ""
Else
    If IsNull(Data1.Recordset("BF_BCODE")) Then
        clpCode(1).Text = ""
    Else
        If glbLinamar Then
            clpCode(1).Text = Mid(Data1.Recordset("BF_BCODE"), 4)
        Else
            clpCode(1).Text = Data1.Recordset("BF_BCODE")
        End If
        If glbWFC Then
            If IsNull(Data1.Recordset("BF_SPOUSE_COMP")) Then
                txtSComp.Text = ""
            Else
                txtSComp.Text = Data1.Recordset("BF_SPOUSE_COMP")
            End If
            If IsNull(Data1.Recordset("BF_SPOUSE_PLAN")) Then
                txtSPlan.Text = ""
            Else
                txtSPlan.Text = Data1.Recordset("BF_SPOUSE_PLAN")
            End If
            
            If IsNull(Data1.Recordset("BF_SPOUSE_CERTIFICATE")) Then
                txtSCert.Text = ""
            Else
                txtSCert.Text = Data1.Recordset("BF_SPOUSE_CERTIFICATE")
            End If
        End If
    End If
End If
If Data3.Recordset.EOF Then
     clpBCODE.Text = ""
Else
    If IsNull(Data3.Recordset("BD_BCODE")) Then
         clpBCODE.Text = ""
    Else
        If glbLinamar Then
             clpBCODE.Text = Mid(Data3.Recordset("BD_BCODE"), 4)
        Else
             clpBCODE.Text = Data3.Recordset("BD_BCODE")
        End If
    End If
End If
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call getCodes
Call Display_Value

End Sub
Private Sub EMP_Releate()
Dim rsEMP As New ADODB.Recordset
Dim SQLQ, xYear
lblDOH = "": lblYear = ""
If glbtermopen Then
    If glbOracle Then
        SQLQ = "SELECT ED_DOH,TERM_DOT,ED_EMPTYPE,ED_BENEFIT_GROUP FROM TERM_HREMP "
        SQLQ = SQLQ & " ,TERM_HRTRMEMP "
        SQLQ = SQLQ & " WHERE TERM_HREMP.TERM_SEQ=TERM_HRTRMEMP.TERM_SEQ"
        SQLQ = SQLQ & " AND TERM_HREMP.TERM_SEQ=" & glbTERM_Seq
    Else
        SQLQ = "SELECT ED_DOH,Term_DOT,ED_EMPTYPE,ED_BENEFIT_GROUP FROM Term_HREMP "
        SQLQ = SQLQ & " INNER JOIN Term_HRTRMEMP "
        SQLQ = SQLQ & " ON Term_HREMP.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ"
        SQLQ = SQLQ & " WHERE Term_HREMP.TERM_SEQ=" & glbTERM_Seq
    End If
    rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
    If Not rsEMP.EOF Then
    If IsDate(rsEMP("ED_DOH")) Then
        lblDOH = rsEMP("ED_DOH")
        If IsDate(rsEMP("Term_DOT")) Then
            xYear = DateDiff("d", CVDate(lblDOH), rsEMP("Term_DOT"))
        End If
    End If
    End If
Else
    rsEMP.Open "SELECT ED_DOH,ED_EMPTYPE,ED_BENEFIT_GROUP FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEMP.EOF Then
    If IsDate(rsEMP("ED_DOH")) Then
        lblDOH = rsEMP("ED_DOH")
        xYear = DateDiff("d", CVDate(lblDOH), Date)
    End If
    End If
End If
If IsNumeric(xYear) Then
    xYear = Round(xYear / 365, 1)
    lblYear = xYear & IIf(xYear <> 1, " Years", " Year")
End If

If Not rsEMP.EOF Then
If IsNull(rsEMP("ED_BENEFIT_GROUP")) Then
    fEBGroup = "NOGROUP"
Else
    fEBGroup = rsEMP("ED_BENEFIT_GROUP")
End If
End If

If glbLinamar Then
    Dim EMPTYPE
    If Not rsEMP.EOF Then
        EMPTYPE = rsEMP("ED_EMPTYPE")
    End If
    If EMPTYPE <> "1" And EMPTYPE <> "3" Then
        fUpdable = False
        cmdRecalAll.Enabled = False
        cmdRecal.Enabled = False
        cmdBens.Enabled = False
    Else
        fUpdable = True
        cmdRecalAll.Enabled = True
        cmdRecal.Enabled = True
        cmdBens.Enabled = True
    End If
End If
End Sub
Private Sub BENCode_Desc()
Dim SQLQ As String, xCode
Dim rsBC As New ADODB.Recordset
On Error GoTo BENCode_Error
fglbDWM = ""
fglbTB_WP = ""
If Len(clpCode(1)) = 0 Then Exit Sub
xCode = Replace(clpCode(1), "'", "''")

SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'BNCD'"
If glbLinamar Then
    SQLQ = SQLQ & " AND (TB_KEY ='" & txtDiv & xCode & "'"
    SQLQ = SQLQ & " OR TB_KEY = 'ALL" & xCode & "') "
Else
    SQLQ = SQLQ & " AND TB_KEY = '" & xCode & "'"
End If
rsBC.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsBC.EOF Then Exit Sub
If Not IsNull(rsBC("TB_USR1")) Then fglbDWM = rsBC("TB_USR1")
If IsNumeric(rsBC("TB_USR2")) Then fglbTB_WP = rsBC("TB_USR2")
rsBC.Close
Exit Sub

BENCode_Error:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "BENCode Snap", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub Display_Value()
Dim SQLQ
Dim rsEmpCert As New ADODB.Recordset
If FrmDetails.Visible = True Then
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    Else
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
          SQLQ = " select * FROM Term_HRBENFT "
          SQLQ = SQLQ & " WHERE BF_BENE_ID = " & Data1.Recordset!BF_BENE_ID
          SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
          rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
          SQLQ = " Select * FROM HRBENFT "
          SQLQ = SQLQ & " WHERE BF_BENE_ID = " & Data1.Recordset!BF_BENE_ID
          SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
          rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        clpGroup.seleEMPCode = fEBGroup
        Call Set_Control("R", Me, rsDATA)
        dlpDate(0) = rsDATA!BF_EDATE
        Call getCodes
    End If
   
Else
    If Data3.Recordset.EOF Or Data3.Recordset.BOF Then
        Call Set_Control3("B", rsDATA3)
        If rsDATA3.State <> 0 Then rsDATA3.Close
        If glbtermopen Then
            rsDATA3.Open Data3.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA3.Open Data3.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        DoEvents
        
        Call SET_UP_MODE2
         Call lblEEID_Change
        Exit Sub
    End If
   
    If rsDATA3.State <> 0 Then rsDATA3.Close
    If glbtermopen Then
      SQLQ = "select * FROM Term_HRBENS "
      SQLQ = SQLQ & " WHERE BD_ID = " & Data3.Recordset!BD_ID
      SQLQ = SQLQ & " ORDER BY BD_BCODE "
      rsDATA3.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
      SQLQ = "select * FROM HRBENS "
      SQLQ = SQLQ & " WHERE BD_ID= " & Data3.Recordset!BD_ID
      SQLQ = SQLQ & " ORDER BY BD_BCODE "
      rsDATA3.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    Call lblEEID_Change
    If rsDATA3.EOF Or rsDATA3.BOF Then Exit Sub
    
    Call Set_Control3("R", rsDATA3)
 End If
 If FrmDetails.Visible = True Then
    Call SET_UP_MODE
 Else
    Call SET_UP_MODE2
 End If
 
 If glbWFC And Not glbtermopen Then 'Ticket #13836
    locManulifeCertNo = ""
    SQLQ = "SELECT ED_EMPNBR, ED_USER_TEXT1 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsEmpCert.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpCert.EOF Then
        If Not IsNull(rsEmpCert("ED_USER_TEXT1")) Then
            If Len(rsEmpCert("ED_USER_TEXT1")) > 0 Then
                locManulifeCertNo = rsEmpCert("ED_USER_TEXT1")
            End If
        End If
    End If
 End If
 Me.cmdModify_Click
 End Sub
 
Private Sub Set_Control3(Act As String, Optional rsTA As ADODB.Recordset)
If Act = "U" Then
    If Len(MedSplitPc) = 0 Then
        rsTA!BD_PC = Null
    Else
        rsTA!BD_PC = MedSplitPc
    End If
    
    If Len(txtBeneName.Text) = 0 Then
        rsTA!BD_BNAME = Null
    Else
        rsTA!BD_BNAME = txtBeneName.Text
    End If
    
    If Len(dlpDate(1).Text) = 0 Then
        rsTA!BD_DOB = Null
    Else
        rsTA!BD_DOB = dlpDate(1).Text
    End If
    
    If dlpDate(3).Visible Then
        If Len(dlpDate(3).Text) = 0 Then
            rsTA!BD_DEATHDATE = Null
        Else
            rsTA!BD_DEATHDATE = dlpDate(3).Text
        End If
    End If
   
    If Len(lblRel) = 0 Then
       rsTA!BD_RELATE = Null
    Else
       rsTA!BD_RELATE = lblRel
    End If
    
    'Ticket #21021 Franks 02/28/2012 - begin
    'for WFC
    If chkSepAgree.Visible Then
        rsTA!BD_SEP_AGREE = chkSepAgree.Value
    End If
    If chkSpouseEnt.Visible Then
        rsTA!BD_SPOUSE_ENT = chkSpouseEnt.Value
    End If
    If dlpSepDate.Visible Then
        If Len(dlpSepDate.Text) = 0 Then
            rsTA!BD_SPE_DATE = Null
        Else
            rsTA!BD_SPE_DATE = dlpSepDate.Text
        End If
    End If
    'Ticket #21021 Franks 02/28/2012 - end
    
    'Ticket #24275 Franks 08/26/2013 - begin
    'for WFC
    If dlpDate(4).Visible Then
        If IsDate(dlpDate(4).Text) Then rsTA!BD_END_DATE = dlpDate(4).Text Else rsTA!BD_END_DATE = Null
    End If
    If clpCode(2).Visible Then
        If Len(clpCode(2).Text) = 0 Then rsTA!BD_REASON = Null Else rsTA!BD_REASON = clpCode(2).Text
    End If
    If clpCode(0).Visible Then
        If Len(clpCode(0).Text) = 0 Then rsTA!BD_PENSIONTYPE = Null Else rsTA!BD_PENSIONTYPE = clpCode(0).Text
    End If
    'Ticket #24275 Franks 08/26/2013 - end
ElseIf Act = "B" Then
    MedSplitPc.Text = ""
    txtBeneName.Text = ""
    dlpDate(1).Text = ""
    If dlpDate(3).Visible Then
        dlpDate(3).Text = ""
    End If
    lblRel = ""
    'Ticket #21021 Franks 02/28/2012 - begin
    'for WFC
    If chkSepAgree.Visible Then
        chkSepAgree.Value = False
    End If
    If chkSpouseEnt.Visible Then
        chkSpouseEnt.Value = False
    End If
    If dlpSepDate.Visible Then
        dlpSepDate.Text = ""
    End If
    'Ticket #21021 Franks 02/28/2012 - end
    'Ticket #24275 Franks 08/26/2013 - begin
    'for WFC
    If dlpDate(4).Visible Then dlpDate(4).Text = ""
    If clpCode(2).Visible Then clpCode(2).Text = ""
    If clpCode(0).Visible Then clpCode(0).Text = ""
    'Ticket #24275 Franks 08/26/2013 - end
ElseIf Act = "R" Then
    MedSplitPc.Text = ""
    txtBeneName.Text = ""
    dlpDate(1).Text = ""
    If dlpDate(3).Visible Then
        dlpDate(3).Text = ""
    End If
    lblRel = ""
    'Ticket #21021 Franks 02/28/2012 - begin
    'for WFC
    If chkSepAgree.Visible Then
        chkSepAgree.Value = False
    End If
    If chkSpouseEnt.Visible Then
        chkSpouseEnt.Value = False
    End If
    If dlpSepDate.Visible Then
        dlpSepDate.Text = ""
    End If
    'Ticket #21021 Franks 02/28/2012 - end
    If rsTA.EOF Or rsTA.BOF Then Exit Sub
    
    If IsNull(rsTA!BD_PC) Then
         MedSplitPc.Text = ""
    Else
         MedSplitPc.Text = rsTA!BD_PC
    End If
    
    If IsNull(rsTA!BD_BNAME) Then
         txtBeneName.Text = ""
    Else
         txtBeneName.Text = rsTA!BD_BNAME
    End If
    
    If IsNull(rsTA!BD_DOB) Then
        dlpDate(1).Text = ""
    Else
        dlpDate(1).Text = rsTA!BD_DOB
    End If
    If dlpDate(3).Visible Then
        If IsNull(rsTA!BD_DEATHDATE) Then
            dlpDate(3).Text = ""
        Else
            dlpDate(3).Text = rsTA!BD_DEATHDATE
        End If
    End If
    'Ticket #21021 Franks 02/28/2012 - begin
    'for WFC
    If chkSepAgree.Visible Then
        If IsNull(rsTA!BD_SEP_AGREE) Then
            chkSepAgree.Value = False
        Else
            chkSepAgree.Value = rsTA!BD_SEP_AGREE
        End If
    End If
    If chkSpouseEnt.Visible Then
        If IsNull(rsTA!BD_SPOUSE_ENT) Then
            chkSpouseEnt.Value = False
        Else
            chkSpouseEnt.Value = rsTA!BD_SPOUSE_ENT
        End If
    End If
    If dlpSepDate.Visible Then
        If IsNull(rsTA!BD_SPE_DATE) Then
            dlpSepDate.Text = ""
        Else
            dlpSepDate.Text = rsTA!BD_SPE_DATE
        End If
    End If
    'Ticket #21021 Franks 02/28/2012 - end
    'Ticket #24275 Franks 08/26/2013 - begin
    'for WFC
    If dlpDate(4).Visible Then
        If IsNull(rsTA!BD_END_DATE) Then dlpDate(4).Text = "" Else dlpDate(4).Text = rsTA!BD_END_DATE
    End If
    If clpCode(2).Visible Then
        If IsNull(rsTA!BD_REASON) Then clpCode(2).Text = "" Else clpCode(2).Text = rsTA!BD_REASON
    End If
    If clpCode(0).Visible Then
        If IsNull(rsTA!BD_PENSIONTYPE) Then clpCode(0).Text = "" Else clpCode(0).Text = rsTA!BD_PENSIONTYPE
    End If
    'Ticket #24275 Franks 08/26/2013 - end
    If IsNull(rsTA!BD_RELATE) Then
        lblRel = ""
    Else
        lblRel = rsTA!BD_RELATE
    End If
End If
End Sub
Private Function GetPayPeriod(EmpNbr)
    Dim SQLQ
    Dim rsSal As New ADODB.Recordset
    SQLQ = "SELECT SH_PAYP "
    If glbtermopen Then
        SQLQ = SQLQ & " from Term_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND TERM_SEQ=" & EmpNbr
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    If Not rsSal.EOF Then
        If Not IsNull(rsSal("SH_PAYP")) Then
            GetPayPeriod = rsSal("SH_PAYP")
        Else
            GetPayPeriod = ""
        End If
    Else
        GetPayPeriod = ""
    End If
    rsSal.Close
End Function
Private Function GetPayrollPension()
Dim SQLQ
Dim rsSal As New ADODB.Recordset
Dim xPenAmt
    xPenAmt = 0
    SQLQ = "SELECT SH_EMPNBR, SH_SALCD, SH_SALARY,SH_WHRS "
    SQLQ = SQLQ & " from HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & glbLEE_ID
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsSal.EOF Then
        If rsSal("SH_SALCD") = "A" Then
            xPenAmt = Round(rsSal("SH_SALARY") / 26, 2)
        Else
            If Not IsNull(rsSal("SH_WHRS")) Then
                xPenAmt = Round((rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52) / 52, 2)
            End If
        End If
    End If
    rsSal.Close
    GetPayrollPension = xPenAmt
End Function
Private Function GetSalCD(EmpNbr)
    Dim SQLQ
    Dim rsSal As New ADODB.Recordset
    SQLQ = "SELECT SH_SALCD "
    If glbtermopen Then
        SQLQ = SQLQ & " from Term_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND TERM_SEQ=" & EmpNbr
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    If Not rsSal.EOF Then
        If Not IsNull(rsSal("SH_SALCD")) Then
            GetSalCD = rsSal("SH_SALCD")
        Else
            GetSalCD = ""
        End If
    Else
        GetSalCD = ""
    End If
    rsSal.Close
End Function

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
If FrmDetails.Visible = True Then
    UpdateRight = gSec_Upd_Benefits
Else
    UpdateRight = gSec_Upd_Beneficiary
End If
End Property

Public Property Get Addable() As Boolean
Addable = fUpdable
End Property

Public Property Get Updateble() As Boolean
Updateble = fUpdable
End Property

Public Property Get Deleteble() As Boolean
Deleteble = fUpdable
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If rsDATA.State = 0 Then
    Exit Sub
End If
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
If Not UpdateRight Then TF = False

Call set_Buttons(UpdateState)
Call ST_UPD_MODE(TF)
End Sub
Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEBENEFITS.Caption = "Benefits / Beneficiaries - " & Left$(glbLEE_SName, 5)
    frmEBENEFITS.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If

If glbWFC Then 'Ticket #24275 Franks 08/28/2013
    If glbSkip Then
        If FrmBens.Visible Then
            If WFCisEmpNoChanged Then
                Call WFCPenTypeList
                If Not WFCIsAllPenTypeCurrent Then
                    MsgBox "This employee does not have a beneficiary assigned to all pension types. Please add the beneficiary for the missing Pension Type(s)."
                End If
                'Call WFCPenFieldsEnable(False) 'Ticket #24317 Franks 09/17/2013
            End If
            If glbtermopen Then locWFCEmpID = glbTERM_Seq Else locWFCEmpID = glbLEE_ID
        End If
    End If
End If
    
End Sub

Public Sub SET_UP_MODE2()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf Data3.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
If Not UpdateRight Then TF = False
Call set_Buttons(UpdateState)
Call ST_UPD_MODE(TF)

End Sub

Function isChangedBens()

Dim tmDat As New ADODB.Recordset

Dim x%, SQLQ

'SQLQ = "SELECT BD_EMPNBR, BD_BCODE, BD_BNAME, BD_RELATE, BD_DOB, BD_PC "
'SQLQ = SQLQ & " FROM HRBENS WHERE BD_BCODE = '" & clpBCODE.Text & "' AND BD_EMPNBR = " & glbLEE_ID
'tmDat.Open SQLQ, gdbAdoIhr001, adOpenDynamic
isChangedBens = False
If fglbNew Then isChangedBens = True: Exit Function
If Data1.Recordset.EOF Then Exit Function

If Data3.Recordset.EOF Or Data3.Recordset.BOF Then
    Exit Function
End If

If clpBCODE.Text <> Data3.Recordset("BD_BCODE") Then GoTo chk
If txtBeneName.Text <> Data3.Recordset("BD_BNAME") Then GoTo chk
If lblRel <> Data3.Recordset("BD_RELATE") Then GoTo chk
If dlpDate(1).Text <> Data3.Recordset("BD_DOB") Then GoTo chk
If MedSplitPc.Text <> Data3.Recordset("BD_PC") Then GoTo chk

Exit Function
chk:
    isChangedBens = True
End Function

Private Sub Set_Group_Benefit()
    Dim rsGBEN As New ADODB.Recordset
    Dim xBenList
    Dim SQLQ
'    clpGroup.seleEMPCode = fEBGroup
    If Len(clpGroup.Text) = 0 Then
        clpCode(1).seleEMPCode = ""
    Else
        SQLQ = "SELECT BM_BCODE FROM HR_BENEFITS_GROUP "
        SQLQ = SQLQ & " WHERE BM_BENEFIT_GROUP='" & Trim(clpGroup.Text) & "'"
        rsGBEN.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        xBenList = ""
        Do Until rsGBEN.EOF
            xBenList = xBenList & rsGBEN("BM_BCODE") & ","
            rsGBEN.MoveNext
        Loop
        If xBenList <> "" Then xBenList = Left(xBenList, Len(xBenList) - 1)
        clpCode(1).seleEMPCode = xBenList
        rsGBEN.Close
    End If
End Sub

Private Sub ResetValues(Optional xNoCovType = "N")
If clpCode(1).Text = "" Or clpGroup.Text = "" Then Exit Sub

Dim rsBGMST As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim SQLQ As String
Dim xACT
Dim xCode As String, xCover As String
Dim xDATE
Dim xPER
Dim xDateAge65

SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & clpGroup & "'"
SQLQ = SQLQ & " AND BM_BCODE='" & clpCode(1).Text & "'"
If xNoCovType = "Y" Then 'Ticket #22411 Franks 08/08/2012
Else
    If Len(txtCovType) = 0 Then
         SQLQ = SQLQ & " AND (BM_COVER='' OR BM_COVER IS NULL) "
    Else
         SQLQ = SQLQ & " AND BM_COVER='" & txtCovType & "'"
    End If
End If
rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

If Not rsBGMST.EOF Then
    If IsDate(rsBGMST("BM_EDATE")) Then
        dlpDate(0).Text = rsBGMST("BM_EDATE")
    Else
        'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
        If glbCompSerial = "S/N - 2420W" And clpCode(1).Text = "PEN" Then
            dlpDate(0).Text = CountEDate(glbLEE_ID, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), , , clpCode(1).Text)
        Else
            dlpDate(0).Text = CountEDate(glbLEE_ID, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"))
        End If
    End If
    If IsNull(rsBGMST("BM_COVER")) Then txtCovType = "" Else txtCovType = rsBGMST("BM_COVER")
    medCovAmount = rsBGMST("BM_AMT")
    medPayPeriodAmount = rsBGMST("BM_PPAMT")
    medUnitCost = rsBGMST("BM_UNITCOST")
    medPPE = rsBGMST("BM_PCE")
    medPPComp = rsBGMST("BM_PCC")
    medEECost = rsBGMST("BM_ECOST")
    medCompCost = rsBGMST("BM_CCOST")
    medTCost = rsBGMST("BM_TCOST")
    medMaxAmnt = rsBGMST("BM_MAXDOL")
    If Not IsNull(rsBGMST("BM_PREMIUM")) Then
        lblAP = rsBGMST("BM_PREMIUM")
    End If
    
    txtPer = rsBGMST("BM_PER")
    medMCCOST = rsBGMST("BM_MTHCCOST")
    medMECOST = rsBGMST("BM_MTHECOST")
    If IsNull(rsBGMST("BM_TAXBEN")) Then txtTAXBEN = "" Else txtTAXBEN = rsBGMST("BM_TAXBEN")
    If IsNull(rsBGMST("BM_SALARYDEPENDANT")) Then txtSalDepn = "" Else txtSalDepn = rsBGMST("BM_SALARYDEPENDANT")
    If IsNull(rsBGMST("BM_MINIMUM")) Then medMinCover = "" Else medMinCover = rsBGMST("BM_MINIMUM")
    If IsNull(rsBGMST("BM_FACTOR")) Then medSalFactor = "" Else medSalFactor = rsBGMST("BM_FACTOR")
    If IsNull(rsBGMST("BM_ROUND")) Then txtRoundFactor = "0" Else txtRoundFactor = rsBGMST("BM_ROUND")
    If IsNull(rsBGMST("BM_MAXIMUM")) Then medMaxCover = "" Else medMaxCover = rsBGMST("BM_MAXIMUM")
    If IsNull(rsBGMST("BM_NEXTNEAREST")) Then lblRound = "" Else lblRound = rsBGMST("BM_NEXTNEAREST")
    
    If IsNull(rsBGMST("BM_WAITPERIOD")) Then txtWaitPeriod = "" Else txtWaitPeriod = rsBGMST("BM_WAITPERIOD")
    If IsNull(rsBGMST("BM_DWM")) Then txtDWM = "" Else txtDWM = rsBGMST("BM_DWM")
    If IsNull(rsBGMST("BM_COMMENTS")) Then memComments = "" Else memComments = rsBGMST("BM_COMMENTS")
    If IsNull(rsBGMST("BM_PTAX")) Then txtPreAftTax = "" Else txtPreAftTax = rsBGMST("BM_PTAX")
    'Ticket #16286
    If Len(txtPolicy.Text) = 0 Then
        If IsNull(rsBGMST("BM_POLICY")) Then txtPolicy.Text = "" Else txtPolicy.Text = rsBGMST("BM_POLICY")
    End If
    
    If glbCompSerial = "S/N - 2347W" Then
        Dim FTE, USDate
        FTE = GetJHData(glbLEE_ID, "JH_FTENUM", 1)
        USDate = GetEmpData(glbLEE_ID, "ED_USRDAT1", Null)
        If IsNull(USDate) Then dlpDate(0).Text = "" Else dlpDate(0).Text = USDate
        
        If FTE < 1 Then
            If rsBGMST("BM_PCC") = 1 Then
                medPPComp = FTE
                medPPE = 1 - FTE
                medCompCost = FTE * medTCost
                medEECost = (1 - FTE) * medTCost
                medMCCOST = medCompCost / 12
                medMECOST = medEECost / 12
            End If
        End If
    End If
        
    'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
    If glbCompSerial = "S/N - 2290W" And (clpCode(1) = "LTD" And ((clpGroup <> "PARTNERS" And clpGroup <> "ART") Or (clpGroup = ""))) Then 'And cmbDWM.ListIndex >= 0 Then
        'xPER = Left(cmbDWM, 1)
        'If xPER = "W" Then xPER = "ww"
        
        'Get the date for Age 65 or 67 based on the Benefit Group
        If (clpGroup = "") Then
            xDateAge65 = DateAdd("yyyy", 67, CVDate(GetEmpData(glbLEE_ID, "ED_DOB")))
        Else
            xDateAge65 = DateAdd("yyyy", 65, CVDate(GetEmpData(glbLEE_ID, "ED_DOB")))
        End If
        
        'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
        'dlpDate(2).Text = MonthLastDate(DateAdd(xPER, 0 - Val(txtWaitPeriod), CVDate(xDateAge65)))
        'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
        If (clpGroup = "") Then
            dlpDate(2).Text = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
        Else
            dlpDate(2).Text = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
        End If
    End If
    
End If
Call Set_SalCover
  

End Sub

Private Sub Recal_Screen_Values()
    
If Len(Trim(medMinCover)) = 0 Then medMinCover = 0
If Len(Trim(medMaxCover)) = 0 Then medMaxCover = 0
If Len(Trim(txtPer)) = 0 Then txtPer = 0
If Len(medMCCOST) = 0 Or medMCCOST = "." Then medMCCOST = 0
If Len(medMECOST) = 0 Or medMECOST = "." Then medMECOST = 0
If Not glbCElgin Then
    Call Set_SalCover
Else
    Call CalElgin
End If

End Sub


Private Function getProcessDate()
Dim rsEMP As New ADODB.Recordset

getProcessDate = Date
If glbtermopen Then Exit Function

rsEMP.Open "SELECT ED_UNION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
If Not rsEMP.EOF Then
    If IsDate(rsEMP("ED_UNION")) Then
        If rsEMP("ED_UNION") > Date Then
            getProcessDate = rsEMP("ED_UNION")
        End If
    End If
End If
End Function

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        

        'Ticket #18090 - begin
        xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
        End If
        'Ticket #18090 - end
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONBENEFIT")
            'frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice"
            'Ticket #18578
            'frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice - " & lblEEName.Caption
            'Ticket #18755
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR New Benefit Notice - " & xBranch & lblEEName.Caption
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
    'If Err.Number = 364 Then
    '    Exit Sub
    'End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    'Resume Next
    Exit Sub

End Sub

Public Sub imgEmail_Click(Optional xType As String)
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
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
            If Not IsMissing(xType) Then
                If xType = "DELETE" Then
                    frmSendEmail.txtSubject.Text = "info:HR Benefit Delete Notice - " & lblEEName.Caption
                ElseIf xType = "UPDATE" Then
                    frmSendEmail.txtSubject.Text = "info:HR Benefit Update Notice - " & lblEEName.Caption
                Else
                    frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice - " & lblEEName.Caption
                End If
            Else
                frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice - " & lblEEName.Caption
            End If
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

Private Sub WFCNewDBBenforSeparation(xEmpNo, BCode, xNewBNAME, xODOB, xOMedSplitPc, xOBRELATE, Optional xTermSEQ = 0)
Dim rslocBen As New ADODB.Recordset
Dim SQLQ As String
    If xTermSEQ > 0 Then
        SQLQ = "SELECT * FROM Term_HRBENS WHERE TERM_SEQ = " & xTermSEQ & " "
    Else
        SQLQ = "SELECT * FROM HRBENS WHERE BD_EMPNBR = " & xEmpNo & " "
    End If
    SQLQ = SQLQ & "AND BD_BCODE = '" & BCode & "' "
    SQLQ = SQLQ & "AND BD_BNAME = '" & xNewBNAME & "' "
    If rslocBen.State <> 0 Then rslocBen.Close
    rslocBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rslocBen.EOF Then
        rslocBen.AddNew
        rslocBen("BD_EMPNBR") = xEmpNo
        rslocBen("BD_BCODE") = BCode
    End If
    rslocBen("BD_BNAME") = xNewBNAME
    If IsDate(xODOB) Then rslocBen("BD_DOB") = CVDate(xODOB)
    If IsNumeric(xOMedSplitPc) Then
        rslocBen("BD_PC") = xOMedSplitPc '/ 100
    End If
    If Len(xOBRELATE) > 0 Then
        rslocBen("BD_RELATE") = xOBRELATE
    End If
    rslocBen("BD_LDATE") = Date
    rslocBen("BD_LTIME") = Time$
    rslocBen("BD_LUSER") = glbUserID
    If xTermSEQ > 0 Then
        rslocBen("TERM_SEQ") = xTermSEQ
    End If
    rslocBen.Update
End Sub


Private Function getBEDateFrom_clpBCODE(xEmpNo, BCode)
Dim rslocBen As New ADODB.Recordset
Dim SQLQ As String
Dim retval
    retval = Date
    SQLQ = "SELECT * FROM HRBENFT "
    SQLQ = SQLQ & "WHERE BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND BF_BCODE = '" & BCode & "' "
    SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE DESC "
    If rslocBen.State <> 0 Then rslocBen.Close
    rslocBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rslocBen.EOF Then
        retval = rslocBen("BF_EDATE")
    End If
    rslocBen.Close
    
    getBEDateFrom_clpBCODE = retval
End Function

Private Function isAnotherSameSIN(xEmpNo)
Dim rsLocEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN
Dim xMsg As String
Dim retval As Boolean

    retval = False

    'find the matching SIN
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HREMP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo
    End If
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xSIN = rsLocEmp("ED_SIN")
        
        SQLQ = "SELECT * FROM HRP_PENSION_BENEFICIARY WHERE PE_SIN = '" & xSIN & "' "
        'If IsSpouse = "Y" Then
            SQLQ = SQLQ & "AND PE_BEN_RELATE = 'Spouse' " 'clpBCODE
        'Else
        '    SQLQ = SQLQ & "AND NOT PE_BEN_RELATE = 'Spouse' "
        'End If
        'SQLQ = SQLQ & "AND BD_BCODE = '" & clpBCODE.Text & "' "
        SQLQ = SQLQ & "AND NOT (PE_CURRENT = 0) "
        SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSERP') "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSUP') "
        SQLQ = SQLQ & "AND NOT (PE_EMPNBR = " & xEmpNo & ") "
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsPen.EOF Then
            If Not (rsLocEmp("ED_SURNAME") = rsPen("PE_SURNAME") And rsLocEmp("ED_FNAME") = rsPen("PE_FNAME")) Then
                xMsg = "Employee " & rsPen("PE_FNAME") & " " & rsPen("PE_SURNAME") & " has a beneficiary with the same SIN as this employee. "
                xMsg = xMsg & Chr(10) & "You cannot have more than one SPOUSE beneficiary with the same SIN. "
                MsgBox xMsg
                retval = True
            End If
        End If
        rsPen.Close
    End If
    isAnotherSameSIN = retval
End Function
Private Function isSpouseExistWithoutEndDate(xEmpNo, xPenType, IsSpouse) 'Ticket #24275 Franks 08/27/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN
Dim retval As Boolean
    retval = False
    
    'Ticket #25051 Franks 02/06/2014 - begin
    'check from info:HR Beneficiary table, if no Spouse Beneficirary then don't check Pension table
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HRBENS "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HRBENS "
        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & xEmpNo
    End If
    SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
    SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' "
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsLocEmp.EOF Then
        GoTo rec_end
    End If
    rsLocEmp.Close
    'Ticket #25051 Franks 02/06/2014 - end
    
    'find the matching SIN
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HREMP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo
    End If
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xSIN = rsLocEmp("ED_SIN")
        
        SQLQ = "SELECT * FROM HRP_PENSION_BENEFICIARY WHERE PE_SIN = '" & xSIN & "' "
        If IsSpouse = "Y" Then
            SQLQ = SQLQ & "AND PE_BEN_RELATE = 'Spouse' " 'clpBCODE
        Else
            SQLQ = SQLQ & "AND NOT PE_BEN_RELATE = 'Spouse' "
        End If
        'SQLQ = SQLQ & "AND BD_BCODE = '" & clpBCODE.Text & "' "
        SQLQ = SQLQ & "AND NOT (PE_CURRENT = 0) "
        If Len(xPenType) > 0 Then
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "
        End If
        SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSERP') "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSUP') "
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsPen.EOF Then
            If IsNull(rsPen("PE_SPOUSE_ENT")) Or rsPen("PE_SPOUSE_ENT") = 0 Then 'PE_SPOUSE_ENT not checked
                retval = True
            End If
        End If
        rsPen.Close
    End If
rec_end:
    isSpouseExistWithoutEndDate = retval

'    retVal = False
'    SQLQ = "SELECT * FROM HRBENS WHERE BD_EMPNBR = " & xEmpNo & " "
'    If IsSpouse = "Y" Then
'        SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' " 'clpBCODE
'    Else
'        SQLQ = SQLQ & "AND NOT BD_RELATE = 'Spouse' "
'    End If
'    SQLQ = SQLQ & "AND BD_BCODE = '" & clpBCODE.Text & "' "
'    SQLQ = SQLQ & "AND BD_DEATHDATE IS NULL "
'    SQLQ = SQLQ & "AND BD_SPE_DATE IS NULL "
'    If Len(xPenType) > 0 Then
'        SQLQ = SQLQ & "AND BD_PENSIONTYPE = '" & xPenType & "' "
'    End If
'    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsTemp.EOF Then
'        retVal = True
'    End If
'    rsTemp.Close
'    isSpouseExistWithoutEndDate = retVal
End Function

Private Function isHCSAExist(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retval As Boolean
    retval = False
    SQLQ = "SELECT * FROM HRBENFT "
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & "AND (BF_BCODE= 'HCSA' OR BF_BCODE= 'HCSA1') "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = True
    End If
    rsTemp.Close
    isHCSAExist = retval
End Function

Private Function WFCMissEndDate4DCPP()
Dim retval As Boolean
'"   If DCPP Pay Period amount is changed, update the Pension Master with the change.
'A change to zero means that the benefit will need a benefit end date
    retval = False
    If clpCode(1).Text = "DCPP" Then
        If Not fglbNew Then
            If Not OPPAMT = medPayPeriodAmount.Text Then
                If medPayPeriodAmount.Text = 0 Then 'changed to 0
                    If Not IsDate(dlpDate(2).Text) Then 'no end date
                        retval = True
                    End If
                End If
            End If
        End If
    End If
    WFCMissEndDate4DCPP = retval
End Function

Private Function WFC_ScreenSetup() 'Ticket #24275 Franks 08/26/2013
    chkSPlan.Visible = True
    chkSPlan.DataField = "BF_COORDINATION"
    txtSComp.DataField = "BF_SPOUSE_COMP"
    txtSPlan.DataField = "BF_SPOUSE_PLAN"
    txtSCert.DataField = "BF_SPOUSE_CERTIFICATE"
    'Ticket #18566 - Death Process
    lblTitle(47).Visible = True
    dlpDate(3).Visible = True
    'Ticket #18566 - end
    
    'Ticket #21021 Franks 02/28/2012 - begin
    'Beneficiaries
    lblTitle(49).Visible = True
    lblTitle(50).Visible = True
    lblTitle(51).Visible = True
    chkSepAgree.Visible = True
    chkSpouseEnt.Visible = True
    dlpSepDate.Visible = True
    'Ticket #21021 Franks 02/28/2012 - end
    
    'Ticket #24275 Franks 08/26/2013 - begin - beneficiary
    lblTitle(52).Visible = True
    lblTitle(53).Visible = True
    lblTitle(54).Visible = True
    dlpDate(4).Visible = True
    clpCode(2).Visible = True
    clpCode(0).Visible = True
    
    'Call WFCPenFieldsEnable(False)
    'Ticket #24275 Franks 08/26/2013 - end
End Function

Private Sub WFCPenFieldsEnable(xFlag As Boolean)  'Ticket #24275 Franks 08/27/2013
If Not glbWFC Then Exit Sub
If clpBCODE.Text = "DB" Then
    lblTitle(17).Enabled = xFlag
    lblTitle(14).Enabled = xFlag
    lblTitle(15).Enabled = xFlag
    'lblTitle(16).Enabled = xFlag
    lblTitle(18).Enabled = xFlag
    clpBCODE.Enabled = xFlag
    txtBeneName.Enabled = xFlag
    comRelation.Enabled = xFlag
    'dlpDate(1).Enabled = xFlag 'Ticket #24422 Franks 10/03/2013
    MedSplitPc.Enabled = xFlag
    
    'lblTitle(52).Enabled = xFlag
    'lblTitle(53).Enabled = xFlag
    'lblTitle(54).Enabled = xFlag
    'dlpDate(4).Enabled = xFlag
    'clpCode(2).Enabled = xFlag
    'clpCode(0).Enabled = xFlag
Else
    lblTitle(17).Enabled = True
    lblTitle(14).Enabled = True
    lblTitle(15).Enabled = True
    lblTitle(16).Enabled = True
    lblTitle(18).Enabled = True
    clpBCODE.Enabled = True
    txtBeneName.Enabled = True
    comRelation.Enabled = True
    dlpDate(1).Enabled = True
    MedSplitPc.Enabled = True
End If

If glbWFC Then 'Ticket #26779 Franks 03/09/2015
    If clpBCODE.Text = "DB" And Not xFlag Then
        lblWFCPenDisp.Visible = True
    Else
        lblWFCPenDisp.Visible = False
    End If
End If

End Sub
Private Sub WFCDelPenBeneficiary(xEmpNo, xBeneName, xPenType) 'Ticket #24275 Franks 08/27/2013
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer

    SQLQ = "DELETE FROM HRP_PENSION_BENEFICIARY "
    If glbtermopen Then
        SQLQ = SQLQ & " WHERE PE_EMPNBR = " & glbTERM_ID & " "
    Else
        SQLQ = SQLQ & " WHERE PE_EMPNBR = " & xEmpNo & " "
    End If
    If Len(xBeneName) = 0 Then
        SQLQ = SQLQ & "AND (PE_BEN_NAME IS NULL OR PE_BEN_NAME = '') "
    Else
        SQLQ = SQLQ & "AND PE_BEN_NAME = '" & xBeneName & "' "
    End If
    
    If Len(xPenType) > 0 Then
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "
    End If
    gdbAdoIhr001.Execute SQLQ, I

End Sub
Private Function IsEligibleSpouseExist(xEmpNo, xBeneName)  'Ticket #24275 Franks 08/27/2013
Dim rsTBens As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
Dim xRetVal As Boolean
    xRetVal = False
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HRBENS "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HRBENS "
        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & xEmpNo
    End If
    SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
    SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' "
    SQLQ = SQLQ & "AND BD_DEATHDATE IS NULL "
    SQLQ = SQLQ & "AND BD_SPE_DATE IS NULL "
    If rsTBens.State <> 0 Then rsTBens.Close
    rsTBens.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsTBens.EOF Then
        xRetVal = True
    End If
    rsTBens.Close
    IsEligibleSpouseExist = xRetVal
End Function
Private Function WFCEligibleSpouseFields() 'Ticket #24275 Franks 08/29/2013
Dim rsTBens As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
Dim xRetVal As Boolean
    xRetVal = False
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HRBENS "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HRBENS "
        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
    SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' "
    'SQLQ = SQLQ & "AND BD_DEATHDATE IS NULL "
    'SQLQ = SQLQ & "AND BD_SPE_DATE IS NULL "
    If rsTBens.State <> 0 Then rsTBens.Close
    rsTBens.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsTBens.EOF Then
        If Not IsNull(rsTBens("BD_BNAME")) Then txtBeneName.Text = rsTBens("BD_BNAME")
        If Not IsNull(rsTBens("BD_DOB")) Then dlpDate(1).Text = rsTBens("BD_DOB")
        If Not IsNull(rsTBens("BD_PC")) Then MedSplitPc.Text = rsTBens("BD_PC")
    End If
    rsTBens.Close
End Function
Private Function getPensionTypes(xEmpNo)   'Ticket #24275 Franks 08/27/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN
Dim I As Integer
Dim retval
    retval = ""
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HREMP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo
    End If
    'SQLQ = SQLQ & "AND ED_EMPTYPE = 'Y' "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xSIN = rsLocEmp("ED_SIN")
        SQLQ = "SELECT DISTINCT PE_SIN,PE_PENSIONTYPE FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSERP') "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSUP') "
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
        I = 0
        Do While Not rsPen.EOF
            If I = 0 Then
                retval = retval & "'" & rsPen("PE_PENSIONTYPE") & "'"
            Else
                retval = retval & ",'" & rsPen("PE_PENSIONTYPE") & "'"
            End If
            I = I + 1
            rsPen.MoveNext
        Loop
        rsPen.Close
    End If
    rsLocEmp.Close
    getPensionTypes = retval
End Function
Private Function WFCIsPenEligilbe() 'Ticket #24275 Franks 08/28/2013
Dim rsLocEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN
Dim I As Integer
Dim retval
    retval = False
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HREMP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & "AND ED_EMPTYPE = 'Y' "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        retval = True
    End If
    rsLocEmp.Close
    WFCIsPenEligilbe = retval
End Function
Private Function WFCIsAllPenTypeCurrent() 'Ticket #24275 Franks 08/28/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim rsPeBen As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN
Dim I As Integer
Dim retval
    retval = True
    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HREMP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & "AND ED_EMPTYPE = 'Y' "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xSIN = rsLocEmp("ED_SIN")
        SQLQ = "SELECT DISTINCT PE_SIN,PE_PENSIONTYPE FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSERP') "
        SQLQ = SQLQ & "AND NOT (PE_PENSIONTYPE = 'DBSUP') "
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
        I = 0
        Do While Not rsPen.EOF
            'check beneficiay
            SQLQ = "SELECT PE_EMPNBR FROM HRP_PENSION_BENEFICIARY WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & rsPen("PE_PENSIONTYPE") & "' "
            SQLQ = SQLQ & "AND NOT PE_CURRENT = 0"
            If rsPeBen.State <> 0 Then rsPeBen.Close
            rsPeBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsPeBen.EOF Then
                retval = False
            End If
            rsPen.MoveNext
        Loop
        rsPen.Close
    End If
    rsLocEmp.Close
    WFCIsAllPenTypeCurrent = retval

End Function
Private Sub WFCMultiPenType(xPenType, xBenName)  'Ticket #24275 Franks 08/27/2013
Dim rsTBens As New ADODB.Recordset
Dim rsTBenAdd As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim xPeType2
    If Len(xPenTypeList) > 0 Then
        If InStr(1, xPenTypeList, ",") = 0 Then 'one Pension Type
            'do nothing
        Else 'more than one Pension Type
            'o   If the new beneficiary is not SPOUSE, display a message saying "Will this be the beneficiary for all Pension Types?".
            xPeType2 = Replace(xPenTypeList, "'" & xPenType & "'", "")
            xPeType2 = Replace(xPeType2, "'", "")
            xPeType2 = Replace(xPeType2, ",", "")
            'check if xPeType2 is valid Pension Type
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'EPTY' AND  TB_KEY ='" & xPeType2 & "' "
            rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTABL.EOF Then 'found
                
                If glbtermopen Then
                    SQLQ = " SELECT * FROM Term_HRBENS "
                    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
                Else
                    SQLQ = " SELECT * FROM HRBENS "
                    SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
                End If
                SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
                SQLQ = SQLQ & "AND BD_BNAME = '" & xBenName & "' "
                SQLQ = SQLQ & "AND BD_PENSIONTYPE = '" & xPenType & "' "
                If rsTBens.State <> 0 Then rsTBens.Close
                rsTBens.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsTBens.EOF Then
                    'check if there is xPeType2 ihr beneficiary, if not then add it for both
                    If glbtermopen Then
                        SQLQ = " SELECT * FROM Term_HRBENS "
                        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
                    Else
                        SQLQ = " SELECT * FROM HRBENS "
                        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
                    End If
                    SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
                    SQLQ = SQLQ & "AND BD_BNAME = '" & xBenName & "' "
                    SQLQ = SQLQ & "AND BD_PENSIONTYPE = '" & xPeType2 & "' "
                    If rsTBenAdd.State <> 0 Then rsTBenAdd.Close
                    rsTBenAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsTBenAdd.EOF Then
                        rsTBenAdd.AddNew
                        rsTBenAdd("BD_PENSIONTYPE") = xPeType2
                        rsTBenAdd("BD_EMPNBR") = rsTBens("BD_EMPNBR")
                        rsTBenAdd("BD_BCODE") = rsTBens("BD_BCODE")
                        rsTBenAdd("BD_BNAME") = rsTBens("BD_BNAME")
                        rsTBenAdd("BD_RELATE") = rsTBens("BD_RELATE")
                        rsTBenAdd("BD_DOB") = rsTBens("BD_DOB")
                        rsTBenAdd("BD_PC") = rsTBens("BD_PC")
                        rsTBenAdd("BD_LDATE") = Date
                        rsTBenAdd("BD_LTIME") = Time$
                        rsTBenAdd("BD_LUSER") = glbUserID
                        rsTBenAdd("BD_PAYROLL_ID") = rsTBens("BD_PAYROLL_ID")
                        rsTBenAdd("BD_DEATHDATE") = rsTBens("BD_DEATHDATE")
                        rsTBenAdd("BD_SEP_AGREE") = rsTBens("BD_SEP_AGREE")
                        rsTBenAdd("BD_SPOUSE_ENT") = rsTBens("BD_SPOUSE_ENT")
                        rsTBenAdd("BD_SPE_DATE") = rsTBens("BD_SPE_DATE")
                        rsTBenAdd("BD_END_DATE") = rsTBens("BD_END_DATE")
                        rsTBenAdd("BD_REASON") = rsTBens("BD_REASON")
                        rsTBenAdd.Update
                        DoEvents
                        
                        If glbtermopen Then
                            Call WFCPensionBeneficiary(glbTERM_ID, "DB", glbTERM_Seq, "NEW", , , , xPeType2, txtBeneName.Text)
                        Else
                            Call WFCPensionBeneficiary(glbLEE_ID, "DB", , "NEW", , , , xPeType2, txtBeneName.Text)
                        End If
                
                    End If
                    rsTBenAdd.Close
                    
                End If
            End If
            rsTABL.Close
        End If
    End If
End Sub
Private Sub WFCMultiBeneficiares() 'Ticket #24275 Franks 08/30/2013
'"   There are 2 current beneficiaries for the same pension type. This cannot happen.
'In HRBENS, there are 2 DB Benefit Codes with no End,  Separation, Death dates.
'The only way to determine current beneficiaries is to check percentage. If equal to 100%,
'make that beneficiary current and add today's date as the End Date of the other beneficiary with a reason code of SYS - System Updated.
Dim rsTBens As New ADODB.Recordset
Dim SQLQ As String
    
    'for new record only
    If Not fglbNew Then Exit Sub

    If glbtermopen Then
        SQLQ = " SELECT * FROM Term_HRBENS "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = " SELECT * FROM HRBENS "
        SQLQ = SQLQ & " WHERE BD_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
    SQLQ = SQLQ & " AND BD_PC = 1 "
    SQLQ = SQLQ & " AND BD_SPE_DATE IS NULL AND BD_DEATHDATE IS NULL AND BD_END_DATE IS NULL " 'Ticket #24337 Franks 09/30/2013
    'SQLQ = SQLQ & "AND BD_BNAME = '" & xBenName & "' "
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & "AND BD_PENSIONTYPE = '" & clpCode(0).Text & "' "
    If rsTBens.State <> 0 Then rsTBens.Close
    rsTBens.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsTBens.EOF Then
        If Len(dlpDate(4).Text) = 0 Then dlpDate(4).Text = Date
        If Len(clpCode(2).Text) = 0 Then clpCode(2).Text = "SYS"
        xBensChg = True
        xCurrentBens = False
        xAsOfBens = dlpDate(4).Text
    End If
    rsTBens.Close
End Sub
Private Sub WFCPenTypeList() 'Ticket #24275 Franks 08/27/2013
    If glbWFC Then
        xPenTypeList = getPensionTypes(glbLEE_ID)
        If Len(xPenTypeList) > 0 Then
            If InStr(1, xPenTypeList, ",") = 0 Then 'one Pension Type
                clpCode(0).Text = Replace(xPenTypeList, "'", "")
            Else 'more than one Pension Type
            End If
            clpCode(0).TransDiv = xPenTypeList
            clpPenType.TransDiv = xPenTypeList 'Ticket #24451 Franks 10/08/2013
        End If
    End If
End Sub

Private Function WFCisEmpNoChanged()
Dim retval As Boolean
    retval = False
    If glbtermopen Then
        If Not locWFCEmpID = glbTERM_Seq Then
            retval = True
        End If
    Else
        If Not locWFCEmpID = glbLEE_ID Then
            retval = True
        End If
    End If
    WFCisEmpNoChanged = retval
End Function
 
Private Sub WFCCCBenEndToGTLD() 'Ticket #25307 Franks 04/09/2014
Dim rsBenT As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
    If glbtermopen Then Exit Sub
    If Not IsDate(dlpDate(2).Text) Then
        Exit Sub 'No End Date then skip
    End If
    If dlpDate(2).Text = OBenEndDate Then
        Exit Sub 'No End Date change then skip
    End If
    'update the same Benefit End Date for GTLD
    SQLQ = "SELECT * FROM  HRBENFT "
    SQLQ = SQLQ & "WHERE BF_EMPNBR = " & glbLEE_ID & " AND BF_BCODE = 'GTLD'"
    If rsBenT.State <> 0 Then rsBenT.Close
    rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBenT.EOF Then
        rsBenT("BF_CEASEDATE") = dlpDate(2).Text
        rsBenT.Update
        Call WFC_AUDITBEN_ByField(glbLEE_ID, "M", "BF_CEASEDATE", rsBenT)
    End If
    rsBenT.Close
End Sub

Private Sub Benefit_Age65AutomaticReduction(Optional xEmpnbr)
    Dim rsBen As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDateAge65 As Date
    Dim xDateAge66 As Date
    Dim flgAge65 As Boolean
    Dim flgAge65plus  As Boolean
    Dim birthdate
    Dim CurrentAge As Double
    
    If glbtermopen Then Exit Sub
    
    'Check if employee is Age 65
    flgAge65 = False
    flgAge65plus = False
    
    If IsMissing(xEmpnbr) Then
        SQLQ = "SELECT ED_EMPNBR, ED_DOB, ED_BENEFIT_GROUP FROM HREMP WHERE ED_EMPNBR IN "
        SQLQ = SQLQ & " (SELECT BF_EMPNBR FROM HRBENFT WHERE BF_BCODE IN ('LIFE10','LIFE3','LIFE5'))"
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_DOB, ED_BENEFIT_GROUP FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT BF_EMPNBR FROM HRBENFT WHERE BF_BCODE IN ('LIFE10','LIFE3','LIFE5'))"
    End If
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        'Get the date for Age 65
        xDateAge65 = DateAdd("yyyy", 65, CVDate(rsHREmp("ED_DOB")))
        
        'Compute Age 66 for comparison
        'xDateAge66 = DateAdd("yyyy", 65, CVDate(rsHREmp("ED_DOB")))
        
        birthdate = CVDate(rsHREmp("ED_DOB"))
        
        'Check if employee already 65 years old but < 66 years
        'flgAge65 = IIf(CVDate(Format(Now, "mm/dd/yyyy")) >= CVDate(xDateAge65) And CVDate(Format(Now, "mm/dd/yyyy")) < CVDate(xDateAge66), True, False)
        
        'Calculate current Age
        If IsDate(rsHREmp("ED_DOB")) Then
            birthdate = CVDate(rsHREmp("ED_DOB"))
            
            CurrentAge = DateDiff("m", birthdate, Now)
            If month(birthdate) = month(Now) Then
                If Day(Now) < Day(birthdate) Then
                    CurrentAge = CurrentAge - 1
                End If
            End If
            CurrentAge = CDbl(CurrentAge / 12)
            
            '65 or 65+ ?
            If CurrentAge >= 65 And CurrentAge < 66 Then
                flgAge65 = True
                flgAge65plus = True
            ElseIf CurrentAge >= 66 Then
                flgAge65 = False
                flgAge65plus = True
            Else
                flgAge65 = False
                flgAge65plus = False
            End If
            
            'Format(Age, "#0.0")
        End If
        
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    'Update Benefit records if PARTNER employee turned 65
    If flgAge65 Then
        'LIFE10 and PARTNERS
        'Reset employee's LIFE10 Benefit if employee has turned 65
        SQLQ = "SELECT * FROM HRBENFT"
        'SQLQ = SQLQ & " WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND BF_GROUP = 'PARTNERS' "
        SQLQ = SQLQ & " WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND (BF_GROUP = '' OR BF_GROUP IS NULL)"
        SQLQ = SQLQ & " AND BF_BCODE = 'LIFE10'"
        'SQLQ = SQLQ & " AND BF_SALARYDEPENDANT = 'Y'"
        rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsBen.EOF Then
            If rsBen("BF_AMT") <> 700000 Then
                'Reset values to match 65 years entitlement
                'rsBen("BF_GROUP") = ""
                rsBen("BF_SALARYDEPENDANT") = "N"
                rsBen("BF_AMT") = 700000      'Fixed Amount
                rsBen("BF_LDATE") = Format(Now, "SHORT DATE")
                rsBen("BF_LUSER") = glbUserID
                rsBen("BF_LTIME") = Time$
                
                'Turn on Age65+
                rsBen("BF_AGE65PLUS") = True
                
                'Shaundra wants the Effective to be updated as well whenever Coverage Amount changes
                rsBen("BF_EDATE") = Format(Now, "SHORT DATE")
                
                rsBen.Update
                
                'Recompute Benefit Cost since Coverage Amount changed
                Call Recalculate_Age65BenefitCost(rsHREmp("ED_EMPNBR"), rsBen("BF_BENE_ID"))
                
                'Audit Update
                If Not AUDITBENF_Age65plus(rsHREmp("ED_EMPNBR"), rsBen("BF_BENE_ID"), "M", 1) Then MsgBox "ERROR - AUDIT FILE"
                
                'Manulife Audit Update - Giving an error - table not found
                'Call AUDIT_MANULIFE_BENF(rsBen("BF_BCODE"), rsBen("BF_EDATE"), rsBen("BF_COVER"), rsBen("BF_POLICY"), rsBen("BF_CEASEDATE"))
            End If
        End If
        rsBen.Close
        Set rsBen = Nothing
    End If
    
    'Update Benefit records if employee turned 65 or is 65+
    If flgAge65 Or flgAge65plus Then
        'LIFE3 or LIFE5 and Not PARTNERS or ART
        'Reset employee's LIFE3 and LIFE5 Benefits if employee has turned 65
        SQLQ = "SELECT * FROM HRBENFT"
        SQLQ = SQLQ & " WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND BF_GROUP <> 'PARTNERS' AND BF_GROUP <> 'ART' "
        SQLQ = SQLQ & " AND (BF_BCODE = 'LIFE3' OR BF_BCODE = 'LIFE5')"
        SQLQ = SQLQ & " AND BF_SALARYDEPENDANT = 'Y'"
        rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsBen.EOF Then
            If rsBen("BF_AMT") <> 0 And Int(CurrentAge) <= 70 Then   'If Coverage is 0 then the employee must be 70
                'Reset values to match 65 years entitlement
                rsBen("BF_GROUP") = ""
                rsBen("BF_AGE65PLUS") = True
                
                rsBen("BF_AMT") = Get_Age65CoverageAmount(rsHREmp("ED_EMPNBR"), CurrentAge, rsBen("BF_AMT"))
                
                'Compute current Age as of on birthday and then add 1 year to compute next Birthday Date.
                If Int(CurrentAge) < 70 Then    'Benefit End at 70
                    rsBen("BF_CEASEDATE") = DateAdd("yyyy", 1, DateAdd("m", (Int(CurrentAge) * 12), CVDate(birthdate)))
                End If
                rsBen("BF_LDATE") = Format(Now, "SHORT DATE")
                rsBen("BF_LUSER") = glbUserID
                rsBen("BF_LTIME") = Time$
                
                'Turn on Age65+
                rsBen("BF_AGE65PLUS") = True
                
                rsBen.Update
                
                'Recompute Benefit Cost based since Coverage Amount changed
                Call Recalculate_Age65BenefitCost(rsHREmp("ED_EMPNBR"), rsBen("BF_BENE_ID"))
                
                'Audit Update
                If Not AUDITBENF_Age65plus(rsHREmp("ED_EMPNBR"), rsBen("BF_BENE_ID"), "M", 1) Then MsgBox "ERROR - AUDIT FILE"
                
                'Manulife Audit Update - Giving an error - table not found
                'Call AUDIT_MANULIFE_BENF(rsBen("BF_BCODE"), rsBen("BF_EDATE"), rsBen("BF_COVER"), rsBen("BF_POLICY"), rsBen("BF_CEASEDATE"))
            End If
        End If
        rsBen.Close
        Set rsBen = Nothing
    
    End If
End Sub

Private Sub Recalculate_Age65BenefitCost(xEmpnbr, xBenID)
    Dim rsBen As New ADODB.Recordset
    Dim SQLQ As String
    
    Dim AMT, tmpTotalCost As Double

    SQLQ = "SELECT * FROM HRBENFT"
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xEmpnbr & " AND BF_BENE_ID = " & xBenID
    rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBen.EOF Then
        If rsBen("BF_TCOST") Then
            tmpTotalCost = rsBen("BF_TCOST")
        End If
        
        AMT = rsBen("BF_AMT")
        
        If rsBen("BF_PREMIUM") = "P" Then
            If Val(rsBen("BF_PER")) > 0 Then tmpTotalCost = (Val(AMT) / Val(rsBen("BF_PER"))) * Val(rsBen("BF_UNITCOST")) Else tmpTotalCost = 0
        End If
        
        rsBen("BF_ECOST") = Val(tmpTotalCost) * Val(rsBen("BF_PCE"))
        rsBen("BF_CCOST") = Val(tmpTotalCost) * Val(rsBen("BF_PCC"))
        
        If glbCompSerial = "S/N - 2262W" Then 'Wellington
            If Right(Round(CStr(Val(rsBen("BF_ECOST")) / 12), 3), 1) = 5 Then
                If rsBen("BF_BCODE") = "5ADB" Or rsBen("BF_BCODE") = "5GRB" Or rsBen("BF_BCODE") = "5LTB" Or _
                    rsBen("BF_BCODE") = "6ADB" Or rsBen("BF_BCODE") = "6GRB" Or rsBen("BF_BCODE") = "6LTB" Or _
                    rsBen("BF_BCODE") = "8GRB" Or rsBen("BF_BCODE") = "4ADW" Or rsBen("BF_BCODE") = "4GRW" Or _
                    rsBen("BF_BCODE") = "4LTW" Or rsBen("BF_BCODE") = "1GRB" Or rsBen("BF_BCODE") = "GRLB" Then
        
                    rsBen("BF_MTHECOST") = Val(rsBen("BF_ECOST")) / 12
                Else
                    rsBen("BF_MTHECOST") = Round((Val(rsBen("BF_ECOST")) / 12) - 0.005, 2)
                End If
            Else
                rsBen("BF_MTHECOST") = Val(rsBen("BF_ECOST")) / 12
            End If
            If Right(Round(CStr(Val(rsBen("BF_CCOST")) / 12), 3), 1) = 5 Then
                If rsBen("BF_BCODE") = "5ADB" Or rsBen("BF_BCODE") = "5GRB" Or rsBen("BF_BCODE") = "5LTB" Or _
                    rsBen("BF_BCODE") = "6ADB" Or rsBen("BF_BCODE") = "6GRB" Or rsBen("BF_BCODE") = "6LTB" Or _
                    rsBen("BF_BCODE") = "8GRB" Or rsBen("BF_BCODE") = "4ADW" Or rsBen("BF_BCODE") = "4GRW" Or _
                    rsBen("BF_BCODE") = "4LTW" Or rsBen("BF_BCODE") = "1GRB" Or rsBen("BF_BCODE") = "GRLB" Then
        
                    rsBen("BF_MTHCCOST") = Val(rsBen("BF_CCOST")) / 12
                Else
                    rsBen("BF_MTHCCOST") = Round((Val(rsBen("BF_CCOST")) / 12) - 0.005, 2)
                End If
            Else
                rsBen("BF_MTHCCOST") = Val(rsBen("BF_CCOST")) / 12
            End If
        
        ' danielk - 12/30/2002 - Priority C changes for 7.0
        ' Franks Sep 4,2002 #2774 for Mitchell Plastic
        ElseIf (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then
            rsBen("BF_MTHECOST") = Val(rsBen("BF_ECOST")) / 12
            rsBen("BF_MTHCCOST") = Val(rsBen("BF_CCOST")) / 12
        Else
            rsBen("BF_MTHECOST") = 0
            rsBen("BF_MTHCCOST") = 0
        End If
        
        'Added by Bryan Aug 21, 2007 - Ticket#13546
        If rsBen("BF_SALARYDEPENDANT") = "Y" Then
            If Val(rsBen("BF_PCC")) + Val(rsBen("BF_PCE")) = 1 Then 'If the total is 100%
                rsBen("BF_TCOST") = tmpTotalCost
            Else 'if the total is not  100%
                rsBen("BF_TCOST") = Val(rsBen("BF_CCOST")) + Val(rsBen("BF_ECOST"))
            End If
        Else
            rsBen("BF_TCOST") = tmpTotalCost
        End If
        
        'Vitalaire
        If glbCompSerial = "S/N - 2380W" Then
            Call CalcPP(Trim(rsBen("BF_BCODE")), Trim(rsBen("BF_GROUP")))
        
            Select Case Trim(rsBen("BF_GROUP"))
            Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"   'Ticket #18963, Ticket #24537 - more codes
                rsBen("BF_PPAMT") = Val(rsBen("BF_ECOST")) / 52
            Case Else
                If (rsBen("BF_PPAMT") = "0" Or rsBen("BF_PPAMT") = "" Or IsNull(rsBen("BF_PPAMT"))) Then
                    rsBen("BF_PPAMT") = Val(rsBen("BF_MTHECOST")) / 2
                End If
            End Select
        End If
        'Franks Sep 4,2002 #2774 for Mitchell Plastic
        
        rsBen.Update
    End If
    rsBen.Close
    Set rsBen = Nothing
End Sub

Private Function Get_Age65CoverageAmount(xEmpnbr, xAge, xDefAmt) As Double
    Dim rsBen As New ADODB.Recordset
    Dim SQLQ As String
    Dim xSalary
    
    'Based on employee's Age, recalculate Coverage Amount.
    'The Coverage Amount starts reducing by 50% from Age 65 upto Age 70. Every year in this period of Age 65 - 70, the
    'Coverage Amount is reduced by additinal 50%.
    
    'The Coverage Amount has been recalculated already based on the standard format and employee's Salary, before this
    'function was called.
    'Retrieve Coverage Amount of the employee and re/apply the 50% based on the employee's Age.
    
    'Salary Dependent Benefit, Age65+, Benefit Group <> PARTNERS and <> ART, Benefit Code LIFE3 or LIFE5
    SQLQ = "SELECT * FROM HRBENFT"
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xEmpnbr & " AND BF_GROUP <> 'PARTNERS' AND BF_GROUP <> 'ART' "
    SQLQ = SQLQ & " AND (BF_BCODE = 'LIFE3' OR BF_BCODE = 'LIFE5')"
    SQLQ = SQLQ & " AND BF_SALARYDEPENDANT = 'Y'"
    'SQLQ = SQLQ & " AND BF_AGE65PLUS <> 0"
    rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBen.EOF Then
        Get_Age65CoverageAmount = xDefAmt
        
        'Age 65?
        If xAge = 65 Then
            '50% reduction
            Get_Age65CoverageAmount = rsBen("BF_AMT") * (50 / 100)
        ElseIf xAge = 66 Then
            '50% reduction
            'Another 50% reduction on the result
            Get_Age65CoverageAmount = (rsBen("BF_AMT") * (50 / 100)) '* (50 / 100)
            
        ElseIf xAge = 67 Then
            '50% reduction
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            Get_Age65CoverageAmount = (rsBen("BF_AMT") * (50 / 100)) '* (50 / 100)) * (50 / 100)
        ElseIf xAge = 68 Then
            '50% reduction
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            Get_Age65CoverageAmount = (rsBen("BF_AMT") * (50 / 100)) '* (50 / 100)) * (50 / 100)) * (50 / 100)
        ElseIf xAge = 69 Then
            '50% reduction
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            'Another 50% reduction on the result
            Get_Age65CoverageAmount = (rsBen("BF_AMT") * (50 / 100)) '* (50 / 100)) * (50 / 100)) * (50 / 100)) * (50 / 100)
        ElseIf xAge = 70 Then
            'Coverage Amount goes down to 0
            Get_Age65CoverageAmount = 0
        End If
        
    End If
    rsBen.Close
    Set rsBen = Nothing

End Function

Private Function AUDITBENF_Age65plus(xEmpnbr, xBenID, ACTX, aType)
Dim rsBen As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsEMP As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim SQLQ As String

On Error GoTo AUDIT_ERR

AUDITBENF_Age65plus = False

SQLQ = "SELECT * FROM HRBENFT"
SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xEmpnbr & " AND BF_BENE_ID = " & xBenID
rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsBen.EOF Then
    'Retrieve employee information
    rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr, gdbAdoIhr001, adOpenKeyset
    
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
    strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
    strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
    strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST,AU_CEASEDATE, "
    strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS "
    
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    xADD = False
    
    'If ACTX = "D" Then GoTo MODUPD
    
    'If aType = 1 Then 'BENEFITS
    '    GoTo MODUPD
    'End If
    
    'GoTo MODNOUPD
    
MODUPD:
    
    'Add Audit
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
    rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
    rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    
    If ACTX = "D" Then
        If aType = 1 Then
            rsTA("AU_BCODE") = rsBen("BF_BCODE")
            rsTA("AU_COVER") = rsBen("BF_COVER")
            rsTA("AU_EDATE") = rsBen("BF_EDATE")
            rsTA("AU_MAXDOL") = rsBen("BF_MAXDOL")
            'Frank 01/29/04, ticket #5521
            If IsNumeric(rsBen("BF_PPAMT")) Then
                rsTA("AU_PPAMT") = rsBen("BF_PPAMT")
            End If
            If IsNumeric(rsBen("BF_MTHCCOST")) Then
                rsTA("AU_MTHCCOST") = rsBen("BF_MTHCCOST")
            End If
            If IsNumeric(rsBen("BF_MTHECOST")) Then
                rsTA("AU_MTHECOST") = rsBen("BF_MTHECOST")
            End If
        End If
    Else
        If aType = 1 Then
            rsTA("AU_MTHCCOST") = rsBen("BF_MTHCCOST")
            rsTA("AU_MTHECOST") = rsBen("BF_MTHECOST")
            rsTA("AU_BCODE") = rsBen("BF_BCODE")
            rsTA("AU_TCOST") = rsBen("BF_TCOST")
            rsTA("AU_PPAMT") = rsBen("BF_PPAMT")
            rsTA("AU_CEASEDATE") = IIf(IsDate(rsBen("BF_CEASEDATE")), rsBen("BF_CEASEDATE"), Null)
            rsTA("AU_BAMT") = rsBen("BF_AMT")
        End If
    End If
    
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpnbr
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = xEmpnbr
    
    If Actn = "A" And FrmDetails.Visible Then
        If CVDate(rsBen("BF_EDATE")) > CVDate(Date) Then
            rsTA("AU_LDATE") = Format(rsBen("BF_EDATE"), "SHORT DATE")
        Else
            rsTA("AU_LDATE") = Date
        End If
    Else
        If CVDate(rsBen("BF_EDATE")) > CVDate(Date) And Not IsDate(dlpDate(2).Text) Then
            rsTA("AU_LDATE") = Format(rsBen("BF_EDATE"), "SHORT DATE")
        Else
            If IsDate(rsBen("BF_CEASEDATE")) Then
                If CVDate(rsBen("BF_CEASEDATE")) > CVDate(Date) Then 'Ticket #14867
                    rsTA("AU_LDATE") = Format(rsBen("BF_CEASEDATE"), "SHORT DATE")
                Else
                    rsTA("AU_LDATE") = Date
                End If
            ElseIf CVDate(rsBen("BF_EDATE")) > CVDate(Date) Then
                rsTA("AU_LDATE") = Format(rsBen("BF_EDATE"), "SHORT DATE")
            Else
                rsTA("AU_LDATE") = Date
            End If
        End If
    End If
    
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = ACTX
    rsTA.Update
    
    If glbWFC Then
        If glbEmpCountry = "CANADA" Then 'Ticket #15818, do not pass benefit to payroll
            Call WFCCNDBeneAuditFlag(xEmpnbr)
        End If
        If glbEmpCountry = "U.S.A." And rsBen("BF_BCODE") = "CC" Then 'Ticket #25307 Franks 04/09/2014
            Call WFCCCBenEndToGTLD
        End If
    End If
End If

MODNOUPD:
AUDITBENF_Age65plus = True

Exit Function

AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT AGE65PLUS", "AUDIT FILE", "UPDATE")

If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function AUDIT_MANULIFE_BENF_Age65Plus(xEmpnbr, xBCode, xBEDate, xBCover, xPolicy, xBenEndDate) 'No AU_CEASEDATE in HRAUDIT, Jerry said we will add it in next release
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDIT_MANULIFE_BENF_Age65Plus = False

If Len(xPolicy) = "" Then
    Exit Function
End If

'BENEFIT End Date
'If IsDate(xBenEndDate) Then
'    If OBenEndDate = "" Then
'        GoTo MODUPD
'    Else
'        If IsDate(OBenEndDate) Then
'            If CVDate(xBenEndDate) <> CVDate(OBenEndDate) Then 'Ticket #15591
'                GoTo MODUPD
'            End If
'        End If
'    End If
'Else
'    If IsDate(OBenEndDate) Then
'        GoTo MODUPD
'    End If
'End If

'GoTo MODNOUPD

MODUPD:

rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    rsTB.Close:    GoTo MODNOUPD
End If
If IsNull(rsTB("ED_USER_TEXT1")) Then 'Certificate #
    rsTB.Close:    GoTo MODNOUPD
Else
    If Len(Trim(rsTB("ED_USER_TEXT1"))) = 0 Then
        rsTB.Close:    GoTo MODNOUPD
    End If
End If

rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

rsTA.AddNew
rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
rsTA("MT_PT_TABL") = "EDPT"
rsTA("MT_TYPE") = "T"
rsTA("MT_BENEFIT") = xBCode
rsTA("MT_EDATE") = xBEDate
If IsDate(xBenEndDate) Then
    rsTA("MT_CEASEDATE") = xBenEndDate
End If
If Len(xBCover) > 0 Then rsTA("MT_COVER") = xBCover
If Len(Trim(xPolicy)) > 0 Then
    rsTA("MT_POLICY_NO") = Trim(xPolicy)
End If
rsTA("MT_COMPNO") = "001"
rsTA("MT_EMPNBR") = xEmpnbr
rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
rsTA("MT_UPLOAD") = "N"
rsTA("MT_LUSER") = glbUserID
If Not IsDate(xBenEndDate) Then
    rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
Else
    If CVDate(xBenEndDate) < CVDate(Date) Then 'WFC Ticket #14867
        rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
    Else
        rsTA("MT_LDATE") = Format(xBenEndDate, "SHORT DATE")
    End If
End If
rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
rsTA("MT_LTIME") = Time$

rsTA.Update

MODNOUPD:
AUDIT_MANULIFE_BENF_Age65Plus = True

Exit Function

AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT AGE65PLUS", "MANULIFE AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

