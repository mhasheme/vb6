VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMCourseCode 
   Caption         =   "Course Code"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCEUCred 
      Appearance      =   0  'Flat
      DataField       =   "ES_CEUCREDIT"
      Height          =   285
      Left            =   9555
      MaxLength       =   5
      TabIndex        =   31
      Tag             =   "11-CEU Credit"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      MaxLength       =   8
      TabIndex        =   33
      Tag             =   "00-Search Code"
      Top             =   7440
      Width           =   1125
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   34
      Tag             =   "00-Search Description"
      Top             =   7440
      Width           =   4410
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5910
      TabIndex        =   35
      Tag             =   "Find specific record"
      Top             =   7440
      Width           =   840
   End
   Begin VB.ComboBox cmbFlwUpDWMY 
      Height          =   315
      ItemData        =   "frmMCourseCode.frx":0000
      Left            =   10080
      List            =   "frmMCourseCode.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   3105
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFlwuUpDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "ES_FLWUP_PRD_DWMY"
      Height          =   285
      Left            =   10740
      TabIndex        =   69
      Top             =   3120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbPrvDWMY 
      Height          =   315
      ItemData        =   "frmMCourseCode.frx":0038
      Left            =   10080
      List            =   "frmMCourseCode.frx":0048
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   3825
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrvDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "ES_PRV_PRD_DWMY"
      Height          =   285
      Left            =   10740
      TabIndex        =   68
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbCurDWMY 
      Height          =   315
      ItemData        =   "frmMCourseCode.frx":0070
      Left            =   10080
      List            =   "frmMCourseCode.frx":0080
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   3465
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCurDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "ES_CUR_PRD_DWMY"
      Height          =   285
      Left            =   11280
      TabIndex        =   67
      Top             =   3480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUnqforPos 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique for each Position"
      DataField       =   "ES_UNIQUE_FOR_POS"
      Height          =   225
      Left            =   7170
      TabIndex        =   21
      Top             =   2670
      Visible         =   0   'False
      Width           =   2570
   End
   Begin VB.CheckBox chkCorponly 
      Caption         =   "Corporate Only"
      DataField       =   "ES_CORPONLY"
      Height          =   225
      Left            =   3240
      TabIndex        =   20
      Top             =   6840
      Width           =   1875
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Active"
      DataField       =   "ES_STATUS"
      Height          =   225
      Left            =   1680
      TabIndex        =   19
      Top             =   6840
      Width           =   1275
   End
   Begin VB.TextBox txtCourseHRS 
      Appearance      =   0  'Flat
      DataField       =   "ES_HOURS"
      Height          =   285
      Left            =   1690
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "11-Number of Scheduled Course Hours "
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      DataField       =   "ES_COMPANYNAME"
      Height          =   285
      Left            =   1690
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "00-Company Name"
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtTrainerName 
      Appearance      =   0  'Flat
      DataField       =   "ES_TRAINNER"
      Height          =   285
      Left            =   1690
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "00-Trainer Name"
      Top             =   4085
      Width           =   3855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8760
      MaxLength       =   25
      TabIndex        =   49
      Text            =   "LUser"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7080
      MaxLength       =   25
      TabIndex        =   48
      Text            =   "LTime"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5400
      MaxLength       =   25
      TabIndex        =   47
      Text            =   "Ldate"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdExport 
         Appearance      =   0  'Flat
         Caption         =   "E&xport"
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
         Left            =   9360
         TabIndex        =   46
         Tag             =   "Print Division Listing"
         Top             =   120
         Width           =   1350
      End
      Begin VB.CommandButton cmdImport 
         Appearance      =   0  'Flat
         Caption         =   "&Import"
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
         Left            =   7980
         TabIndex        =   45
         Tag             =   "Print Division Listing"
         Top             =   120
         Width           =   1350
      End
      Begin VB.CommandButton cmdMissing 
         Appearance      =   0  'Flat
         Caption         =   "P&opulate"
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
         Left            =   6480
         TabIndex        =   44
         Tag             =   "Print Division Listing"
         Top             =   120
         Width           =   1350
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         Left            =   5775
         TabIndex        =   43
         Tag             =   "Print Division Listing"
         Top             =   105
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
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
         Left            =   4959
         TabIndex        =   42
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
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
         Left            =   4145
         TabIndex        =   41
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
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
         Height          =   375
         Left            =   3271
         TabIndex        =   40
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Height          =   375
         Left            =   2457
         TabIndex        =   39
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
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
         Left            =   1643
         TabIndex        =   38
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
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
         Left            =   829
         TabIndex        =   37
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
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
         Left            =   15
         TabIndex        =   36
         Tag             =   "Select this Division"
         Top             =   105
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1935
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Threed.SSPanel spShow 
         Height          =   375
         Left            =   10800
         TabIndex        =   71
         Top             =   120
         Visible         =   0   'False
         Width           =   2715
         _Version        =   65536
         _ExtentX        =   4789
         _ExtentY        =   661
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
         BevelOuter      =   1
         BevelInner      =   2
         FloodType       =   1
         Alignment       =   4
         Autosize        =   3
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmMCourseCode.frx":00A8
      Height          =   2385
      Left            =   120
      OleObjectBlob   =   "frmMCourseCode.frx":00BC
      TabIndex        =   50
      Top             =   120
      Width           =   13335
   End
   Begin INFOHR_Controls.CodeLookup clpEmpCur 
      DataField       =   "ES_EMPCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   8
      Top             =   4770
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CRSCODE"
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Tag             =   "00-Course Code"
      Top             =   2640
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CTYPE"
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Tag             =   "01-Course Type - Code"
      Top             =   3000
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCT"
      MaxLength       =   8
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_TBEMP"
      Height          =   285
      Index           =   0
      Left            =   1690
      TabIndex        =   7
      Tag             =   "20-Amount Employee Paid"
      Top             =   4770
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_OTHER"
      Height          =   285
      Index           =   2
      Left            =   1690
      TabIndex        =   9
      Tag             =   "20-Other Expenses Paid"
      Top             =   5115
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_TBCO"
      Height          =   285
      Index           =   1
      Left            =   1690
      TabIndex        =   11
      Tag             =   "20-Amount Employer Paid"
      Top             =   5460
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_ACCOM"
      Height          =   285
      Index           =   3
      Left            =   1690
      TabIndex        =   13
      Tag             =   "20-Accommodation"
      Top             =   5790
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_COORDINATED"
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   3
      Tag             =   "00-Co-Ordinated By"
      Top             =   3360
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCC"
   End
   Begin INFOHR_Controls.CodeLookup clpOherCur 
      DataField       =   "ES_OTCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   10
      Top             =   5115
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpEmployerCur 
      DataField       =   "ES_EMPLOYCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   12
      Top             =   5460
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpAcomCur 
      DataField       =   "ES_ACOMCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   14
      Top             =   5790
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpTotCur 
      DataField       =   "ES_TOTCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   18
      Top             =   6450
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_LEARNING"
      Height          =   285
      Index           =   4
      Left            =   1690
      TabIndex        =   15
      Tag             =   "20-Accommodation"
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpLearnCur 
      DataField       =   "ES_LEARNINGCUR"
      Height          =   285
      Left            =   2940
      TabIndex        =   16
      Top             =   6120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin MSMask.MaskEdBox medContTotal 
      Height          =   285
      Left            =   1690
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6450
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "Currency"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrvPosRenewal 
      DataField       =   "ES_RENEW_CRS_PRV"
      Height          =   285
      Left            =   9555
      TabIndex        =   26
      Tag             =   "20-Previous Position's Renewal Period"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medFlwUpEffective 
      DataField       =   "ES_RENEW_FOLLOWUP"
      Height          =   285
      Left            =   9555
      TabIndex        =   22
      Tag             =   "20-Follow Up Effective Date Period"
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCurPosRenewal 
      DataField       =   "ES_RENEW_CRS_CUR"
      Height          =   285
      Left            =   9555
      TabIndex        =   24
      Tag             =   "20-Current Position's Renewal Period"
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   9240
      TabIndex        =   28
      Tag             =   "00-Course Code"
      Top             =   5640
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCEUType 
      DataField       =   "ES_CEUTYPE"
      Height          =   285
      Left            =   9240
      TabIndex        =   30
      Top             =   4560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESUT"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_METHODUSED"
      Height          =   285
      Index           =   4
      Left            =   9240
      TabIndex        =   32
      Tag             =   "00-Method Used"
      Top             =   5280
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESMU"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CONDUCT"
      Height          =   285
      Index           =   5
      Left            =   9240
      TabIndex        =   29
      Tag             =   "00-Organization/Individual Instructing - Code"
      Top             =   4200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCB"
   End
   Begin VB.Label lblCEUCred 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEU Credit"
      Height          =   195
      Left            =   7200
      TabIndex        =   76
      Top             =   4965
      Width           =   780
   End
   Begin VB.Label lblShowHSECode 
      Caption         =   "HSE Code Help List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8400
      TabIndex        =   75
      Top             =   7455
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgHelpWFC 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   8040
      Picture         =   "frmMCourseCode.frx":6890
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Conducted By      "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   7200
      TabIndex        =   74
      Top             =   4250
      Width           =   1320
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Method Used"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   7200
      TabIndex        =   73
      Top             =   5325
      Width           =   1605
   End
   Begin VB.Label lblCEUType 
      BackStyle       =   0  'Transparent
      Caption         =   "CEU Type"
      Height          =   195
      Left            =   7200
      TabIndex        =   72
      Top             =   4605
      Width           =   1575
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "WPS Report Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   7200
      TabIndex        =   70
      Top             =   5685
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Follow Up Effective Date Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   7200
      TabIndex        =   66
      Top             =   3165
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Pos. Renewal Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   7200
      TabIndex        =   65
      Top             =   3885
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Pos. Renewal Period"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   7200
      TabIndex        =   64
      Top             =   3525
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   63
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Hours"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   62
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   61
      Top             =   4770
      Width           =   1350
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Expenses $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   60
      Top             =   5115
      Width           =   1665
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employer $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   59
      Top             =   5460
      Width           =   1305
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   58
      Top             =   6450
      Width           =   975
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Accommodation $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   57
      Top             =   5790
      Width           =   1515
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code"
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
      TabIndex        =   56
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   55
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Trainer Name"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   54
      Top             =   4085
      Width           =   1440
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Co-Ordinated By"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   53
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      Caption         =   "Currency"
      Height          =   255
      Index           =   26
      Left            =   2940
      TabIndex        =   52
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Learning Material $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   51
      Top             =   6120
      Width           =   1515
   End
End
Attribute VB_Name = "frmMCourseCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
Dim oCurRen, oPrvRen, oFolRen
Dim oCurRenTyp, oPrvRenTyp, oFolRenTyp
Dim rsDATA As New ADODB.Recordset
Const Excel2007 = 12

Private Sub chkUnqforPos_Click()
    If chkUnqforPos Then
        'lbltitle(12).FontBold = False
        'lbltitle(13).FontBold = False
        lbltitle(14).FontBold = False
        
        medCurPosRenewal.Text = ""
        medPrvPosRenewal.Text = ""
        medFlwUpEffective.Text = ""
        cmbCurDWMY.ListIndex = -1
        cmbPrvDWMY.ListIndex = -1
        cmbFlwUpDWMY.ListIndex = -1
        
        medCurPosRenewal.Enabled = False
        medPrvPosRenewal.Enabled = False
        medFlwUpEffective.Enabled = False
        cmbCurDWMY.Enabled = False
        cmbPrvDWMY.Enabled = False
        cmbFlwUpDWMY.Enabled = False
    Else
        'lbltitle(12).FontBold = True
        'lbltitle(13).FontBold = True
        lbltitle(14).FontBold = True
        medCurPosRenewal.Enabled = True
        medPrvPosRenewal.Enabled = True
        medFlwUpEffective.Enabled = True
        cmbCurDWMY.Enabled = True
        cmbPrvDWMY.Enabled = True
        cmbFlwUpDWMY.Enabled = True
    End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    'If Index = 0 Then Call CrsName_Desc
    If Index = 0 Then Call CourseCode_Type
End Sub

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)

Call modSTUPD(False)  ' reset screen's attributes
cmdClose.SetFocus

Call UpConttotal

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_COURSECODE_MASTER", "Cancel")
Resume Next
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim CrsCode As String, SQLQ As String, Msg$, a%
Dim snapEECRSCodes As New ADODB.Recordset

On Error GoTo DelErr

If Len(clpCode(0)) < 1 Then Exit Sub
CrsCode$ = CStr(clpCode(0))

'SQLQ = "SELECT ES_CRSCODE FROM HREDSEM "
'SQLQ = SQLQ & "WHERE ES_CRSCODE = '" & CrsCode & "'"
'
'If snapEECRSCodes.State <> 0 Then snapEECRSCodes.Close
'snapEECRSCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic
'
'If snapEECRSCodes.BOF And snapEECRSCodes.EOF Then
'    GoTo Lok
'Else
'    Msg$ = lStr("Some mmployees presently assigned to this Code")
'    Msg$ = Msg$ & Chr(10) & "Delete aborted."
'    MsgBox Msg$
'    snapEECRSCodes.Close
'    Exit Sub
'End If

'Lok:    'looks ok to me
'snapEECRSCodes.Close

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

'7.9 - Enhancement - For all clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
    'Chatham-Kent but Chatham-Kent they are not using 7.9
    If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        Msg = Msg & Chr(10) & Chr(10) & "Note: This will also delete from Position's Required Courses and " & Chr(10) & "from the Training List course records of the employees."
    ElseIf glbWFC Then 'Ticket #23317 Franks 02/27/2013
        Msg = Msg & Chr(10) & Chr(10) & "Note: This will also delete from Position's Required Courses, " & Chr(10) & "the Training List course records of the employees."
        Msg = Msg & Chr(10) & "and the Training Development Matrix for HR Stats."
    Else
        Msg = Msg & Chr(10) & Chr(10) & "Note: This will also delete from Position's Required Courses and " & Chr(10) & "from the Training Plan course records of the employees."
    End If
'End If

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


gdbAdoIhr001.BeginTrans

'7.9 - Enhancement - For all clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Call procedure to delete from Position's Required Courses and employee's Training List as well
    Call Deleted_Training_List_Records
'End If

If glbWFC Then 'Ticket #23317 Franks 02/27/2013
    SQLQ = "DELETE FROM WFC_HRST_TRAINING_DEVELOPMENT WHERE TD_CRSCODE = '" & clpCode(0).Text & "'"
    gdbAdoIhr001.Execute SQLQ
End If

rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call UpConttotal

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_COURSECODE_MASTER", "Delete")
Resume Next

End Sub

Private Sub cmdExport_Click()
    Dim rsCourseCodeMst As New ADODB.Recordset
    Dim SQLQ As String
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim xCol, xRow, xRows As Integer
    Dim xlsExpFile As String
    Dim xTrainMatrixPath As String
    Dim xlsFileTmp As String
    Dim i As Integer
    Dim xFDWMY As String
    Dim xCDWMY As String
    Dim Msg As String
    Dim a%
    Dim appVerInt As Double
    
    On Error GoTo Err_CourseCodeMaster_Export
                
    'Get the Export path
    If gsTRAININGMATRIX Then
        xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xTrainMatrixPath) = 0 Then
        xTrainMatrixPath = glbIHRREPORTS
    End If
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "CourseCodeMasterTmp.xls"
    xlsExpFile = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "CourseCodeMaster" & Trim(glbUserID) & ".xls"

    'Msg = "This function will export data from Course Code Master into an Excel spreadsheet 'CourseCodeMaster_" & glbUserID & ".xls'"
    Msg = "This function will export data from Course Code Master into an Excel spreadsheet '" & xlsExpFile & "'"
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do this?"
    a% = MsgBox(Msg, 36, "Confirm Export")
    If a% <> 6 Then Exit Sub

    Screen.MousePointer = HOURGLASS
    
    If Dir(xlsFileTmp) = "" Then
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsExpFile)) <> "" Then Kill xlsExpFile
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsExpFile

    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsExpFile)
    Set exSheet = exBook.Worksheets(1)
        
    
    'Ticket #22166
    appVerInt = Split(exApp.Version, ".")(0)
    If appVerInt - Excel2007 >= 0 Then
        'exApp.ActiveWorkbook.SaveAs (sXLS), 56
        exApp.DisplayAlerts = False
        exBook.SaveAs (xlsExpFile), 56
        exApp.DisplayAlerts = True
    Else
        'exApp.ActiveWorkbook.SaveAs (sXLS), 43
        exApp.DisplayAlerts = False
        exBook.SaveAs (xlsExpFile), 43
        exApp.DisplayAlerts = True
    End If
        
    exSheet.Cells(1, 1) = "Course Code Master Export File"
    exSheet.Cells(2, 1) = "Course Code"
    exSheet.Cells(2, 2) = "Course Type"
    'Ticket #24708 Franks 11/27/2013 - begin
    'exSheet.Cells(2, 3) = lStr("Conducted By")
    exSheet.Cells(2, 3) = lStr("Co-Ordinated By")
    exSheet.Cells(2, 4) = lStr("CEU Type")
    exSheet.Cells(2, 5) = lStr("Method Used")
    'Ticket #24708 Franks 11/27/2013 - end
    exSheet.Cells(2, 3 + 3) = "Unique for each Position"
    If glbWFC Then
        exSheet.Cells(2, 4 + 3) = "WPS Report Code"
        exSheet.Cells(2, 5 + 3) = "Follow Up Effective Date Period"
        exSheet.Cells(2, 6 + 3) = "Follow Up Period in Days/Week/Month/Year"
        exSheet.Cells(2, 7 + 3) = "Current Pos Renewal Period"
        exSheet.Cells(2, 8 + 3) = "Current Renewal in Day/Week/Month/Year"
    Else
        exSheet.Cells(2, 4 + 3) = "Follow Up Effective Date Period"
        exSheet.Cells(2, 5 + 3) = "Follow Up Period in Days/Week/Month/Year"
        exSheet.Cells(2, 6 + 3) = "Current Pos Renewal Period"
        exSheet.Cells(2, 7 + 3) = "Current Renewal in Day/Week/Month/Year"
    End If
        
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER ORDER BY ES_CRSCODE,ES_CTYPE"
    If rsCourseCodeMst.State <> 0 Then rsCourseCodeMst.Close
    rsCourseCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseCodeMst.EOF Then
        xRow = 3
        xRows = rsCourseCodeMst.RecordCount
        i = 0
        If xRows > 0 Then spShow.Visible = True
        
        Do While Not rsCourseCodeMst.EOF
            spShow.FloodPercent = (i / xRows) * 100
            i = i + 1
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsCourseCodeMst("ES_CRSCODE")
            exSheet.Cells(xRow, 2) = rsCourseCodeMst("ES_CTYPE")
            'Ticket #24708 Franks 11/27/2013 - begin
            'exSheet.Cells(xRow, 3) = rsCourseCodeMst("ES_CONDUCT")
            exSheet.Cells(xRow, 3) = rsCourseCodeMst("ES_COORDINATED")
            exSheet.Cells(xRow, 4) = rsCourseCodeMst("ES_CEUTYPE")
            exSheet.Cells(xRow, 5) = rsCourseCodeMst("ES_METHODUSED")
            'Ticket #24708 Franks 11/27/2013 - end
            If rsCourseCodeMst("ES_UNIQUE_FOR_POS") = 1 Then
                exSheet.Cells(xRow, 3 + 3) = "Y"
            Else
                exSheet.Cells(xRow, 3 + 3) = "N"
            End If
            
            xFDWMY = ""
            If rsCourseCodeMst("ES_FLWUP_PRD_DWMY") = "D" Then
                xFDWMY = "Day"
            ElseIf rsCourseCodeMst("ES_FLWUP_PRD_DWMY") = "W" Then
                xFDWMY = "Week"
            ElseIf rsCourseCodeMst("ES_FLWUP_PRD_DWMY") = "M" Then
                xFDWMY = "Month"
            ElseIf rsCourseCodeMst("ES_FLWUP_PRD_DWMY") = "Y" Then
                xFDWMY = "Year"
            End If
            xCDWMY = ""
            If rsCourseCodeMst("ES_CUR_PRD_DWMY") = "D" Then
                xCDWMY = "Day"
            ElseIf rsCourseCodeMst("ES_CUR_PRD_DWMY") = "W" Then
                xCDWMY = "Week"
            ElseIf rsCourseCodeMst("ES_CUR_PRD_DWMY") = "M" Then
                xCDWMY = "Month"
            ElseIf rsCourseCodeMst("ES_CUR_PRD_DWMY") = "Y" Then
                xCDWMY = "Year"
            End If
            
            If glbWFC Then
                exSheet.Cells(xRow, 4 + 3) = rsCourseCodeMst("ES_WPSCODE")
                exSheet.Cells(xRow, 5 + 3) = rsCourseCodeMst("ES_RENEW_FOLLOWUP")
                exSheet.Cells(xRow, 6 + 3) = xFDWMY
                
                exSheet.Cells(xRow, 7 + 3) = rsCourseCodeMst("ES_RENEW_CRS_CUR")
                exSheet.Cells(xRow, 8 + 3) = xCDWMY
            Else
                'exSheet.Cells(xRow, 4) = rsCourseCodeMst("ES_WPSCODE")
                exSheet.Cells(xRow, 4 + 3) = rsCourseCodeMst("ES_RENEW_FOLLOWUP")
                exSheet.Cells(xRow, 5 + 3) = xFDWMY
                
                exSheet.Cells(xRow, 6 + 3) = rsCourseCodeMst("ES_RENEW_CRS_CUR")
                exSheet.Cells(xRow, 7 + 3) = xCDWMY
            End If
            
            xRow = xRow + 1
            
            rsCourseCodeMst.MoveNext
        Loop
    End If
    
    exSheet.Columns.AutoFit
    
    rsCourseCodeMst.Close
    Set rsCourseCodeMst = Nothing
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    spShow.FloodPercent = 100
    spShow.Visible = False
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    If Not LanchXlsW98(xlsExpFile) Then
        Shell "cmd /c " & GetShortName(xlsExpFile)
    End If
    
    If i > 0 Then
        MsgBox "Course Code Master Export complete."
    Else
        MsgBox "No Course Code Master records to export."
    End If
    
Exit Sub

Err_CourseCodeMaster_Export:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HR_COURSECODE_MASTER", "Course Code Master")

Set exSheet = Nothing
Set exBook = Nothing
'exApp.Quit
Set exApp = Nothing
    
End Sub

Private Sub cmdFind_Click()
    Dim SQLQ As String
    
    If Len(txtFindKey) > 0 Then
        SQLQ = "ES_CRSCODE like  '" & txtFindKey.Text & "%'"
        Data1.Recordset.Requery
        Data1.Recordset.Find SQLQ
        If Data1.Recordset.EOF Then
            Data1.Refresh
        Else
            txtFindKey = ""
        End If
        Exit Sub
    End If
    
    If Len(txtFindDesc) > 0 Then
        SQLQ = "COURSEDESC like '" & txtFindDesc.Text & "%'"
        Data1.Recordset.Requery
        Data1.Recordset.Find SQLQ
        If Data1.Recordset.EOF Then
            Data1.Refresh
        Else
            txtFindDesc = ""
        End If
        Exit Sub
    End If
End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImport_Click()
''Ticket #21330 Franks 02/13/2012
''the example file WFC Course Code Master Import File.xls in the folder
''--U:\HR Systems VB6\Custom Features 7x\WFC Custom Programming\Custom Report
'Ticket #21851 Franks 04/19/2012, make Import generic
'the example file "Course Code Master Import File.xls" in the folder
'--U:\HR Systems VB6\Database Files 7.9\MS SQL SERVER

'Ticket #22682 Hemu 02/14/2013 Release 8.0: Made changes and fixes as per requirement in the Release 8.0 document

Dim SQLQ As String, Msg As String, a%
Dim rsCRSCodes As New ADODB.Recordset
Dim rsMaster As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim Inum As Integer
Dim ImportFile
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xCols As Integer, xRows
Dim xCol, xRow
Dim xCode, xType, xWPSCode, xFollowPeriod, xRenewPeriod, xFDWMY, xCDWMY
Dim xUnique
Dim xUptAmt As Integer
Dim CoordinatedBy, xCondBy, xCEUType, xMethodUsed
'Msg = "This function will import data into Course Code Master and will update the Employee's Training Plan "
''Msg = Msg & Chr(10) & "using the 'WFC Course Code Master Import File.xls' file"
'Msg = Msg & Chr(10) & "using the 'Course Code Master Import File.xls' file" & Chr(10) 'Ticket #21851
Msg = "This function will import data from 'Course Code Master Import File.xls' into Course Code Master and "
Msg = Msg & "will also update Employee's Training Plan using Position's Required Courses setup."
Msg = Msg & Chr(10) & Chr(10) & "The file layout is: "
If glbWFC Then
    Msg = Msg & Chr(10) & "Course Code, Course Type, " & lStr("Co-Ordinated By") & ", " & lStr("CEU Type") & ", " & lStr("Method Used") & ", Unique, WPS Report Code, Follow Up Effective Date Period, Follow Up Period in Day/Week/Month/Year, Current Pos Renewal Period, Current Renewal in Day/Week/Month/Year"
Else
    Msg = Msg & Chr(10) & "Course Code, Course Type, " & lStr("Co-Ordinated By") & ", " & lStr("CEU Type") & ", " & lStr("Method Used") & ", Unique, Follow Up Effective Date Period, Follow Up Period in Day/Week/Month/Year, Current Pos Renewal Period, Current Renewal in Day/Week/Month/Year"
End If
Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do this?"
a% = MsgBox(Msg, 36, "Confirm")
If a% <> 6 Then Exit Sub

    'ImportFile = App.Path
    ImportFile = glbIHRREPORTS
    If Right(ImportFile, 1) = "\" Then ImportFile = Left(ImportFile, Len(ImportFile) - 1)
    'ImportFile = ImportFile & "\" & "WFC Course Code Master Import File.xls"
    ImportFile = ImportFile & "\" & "Course Code Master Import File.xls" 'Ticket #21851

    If Dir(ImportFile) = "" Then
        MsgBox ImportFile & " File not Found"
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile) '(AppPath & "benefit group code.xlsx")
    Set exSheet = exBook.Worksheets(1)

    xUptAmt = 0
    xRows = getRowsNspace(exSheet, 2, 1)
    
    If xRows > 3 Then spShow.Visible = True
    DoEvents
    
    For xRow = 3 To xRows
        spShow.FloodPercent = (xRow / xRows) * 100
        DoEvents
        xCode = "": xType = "": xWPSCode = "": xFollowPeriod = "": xRenewPeriod = ""
        
        'Retrive data from spreadsheet
        xCode = exSheet.Cells(xRow, 1)
        xType = exSheet.Cells(xRow, 2)
        
        'Ticket #24708 Franks 11/27/2013 - begin
        'xCondBy = exSheet.Cells(xRow, 3)
        'If IsEmpty(xCondBy) Then xCondBy = ""
        CoordinatedBy = exSheet.Cells(xRow, 3)
        If IsEmpty(CoordinatedBy) Then CoordinatedBy = ""
        xCEUType = exSheet.Cells(xRow, 4)
        If IsEmpty(xCEUType) Then xCEUType = ""
        xMethodUsed = exSheet.Cells(xRow, 5)
        If IsEmpty(xMethodUsed) Then xMethodUsed = ""
        'Ticket #24708 Franks 11/27/2013 - end
        
        xUnique = exSheet.Cells(xRow, 3 + 3) 'Ticket #21851
        If IsEmpty(xUnique) Then xUnique = "N"
        
        If glbWFC Then
            xWPSCode = exSheet.Cells(xRow, 4 + 3)
            xFollowPeriod = exSheet.Cells(xRow, 5 + 3)
            xFDWMY = exSheet.Cells(xRow, 6 + 3)
            xRenewPeriod = exSheet.Cells(xRow, 7 + 3)
            xCDWMY = exSheet.Cells(xRow, 8 + 3)
        Else
            xWPSCode = ""
            xFollowPeriod = exSheet.Cells(xRow, 4 + 3)
            xFDWMY = exSheet.Cells(xRow, 5 + 3)
            xRenewPeriod = exSheet.Cells(xRow, 6 + 3)
            xCDWMY = exSheet.Cells(xRow, 7 + 3)
        End If
        
        'Check if Course Code found in the import file
        If IsEmpty(xCode) Then GoTo NextRec
        
        'Only Import Courses that exists in the HRTABL
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY = '" & xCode & "' "
        If rsCRSCodes.State <> 0 Then rsCRSCodes.Close
        rsCRSCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsCRSCodes.EOF Then
            'Course found in the Table Master
            If IsEmpty(xType) Then
                'If no Course Type found in the Import file then get the Course Type associated with the
                'Course Code in Table Master
                If Not IsNull(rsCRSCodes("TB_USR1")) Then
                    xType = rsCRSCodes("TB_USR1")
                End If
            Else
                If Len(xType) = 0 Then
                    xType = ""
                End If
            End If
            
            'Add or Update the Course in the Course Code Master
            SQLQ = "SELECT * FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & xCode & "' "
            If rsAdd.State <> 0 Then rsAdd.Close
            rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsAdd.EOF Then
                rsAdd.AddNew
                rsAdd("ES_COMPNO") = "001"
                rsAdd("ES_CRSCODE") = rsCRSCodes("TB_KEY")
                rsAdd("ES_STATUS") = 1
                rsAdd("ES_CORPONLY") = 0
            End If
            
            If Len(xType) > 0 Then
                rsAdd("ES_CTYPE") = Left(xType, 8)
            End If
            
            'Ticket #24708 Franks 11/27/2013 - begin
            'If Len(xCondBy) > 0 Then rsAdd("ES_CONDUCT") = Left(xCondBy, 4)
            If Len(CoordinatedBy) > 0 Then rsAdd("ES_COORDINATED") = Left(CoordinatedBy, 4)
            If Len(xCEUType) > 0 Then rsAdd("ES_CEUTYPE") = Left(xCEUType, 8)
            If Len(xMethodUsed) > 0 Then rsAdd("ES_METHODUSED") = Left(xMethodUsed, 4)
            'Ticket #24708 Franks 11/27/2013 - END
            
            If glbWFC Then
                If Not IsEmpty(xWPSCode) Then
                    If Len(xWPSCode) > 0 Then
                        rsAdd("ES_WPSCODE") = Left(xWPSCode, 8)
                    End If
                End If
            End If
            'rsAdd("ES_STATUS") = 1
            'rsAdd("ES_CORPONLY") = 0
            
            'rsAdd("ES_RENEW_FOLLOWUP") = 99
            If Not IsEmpty(xFollowPeriod) Then
                If IsNumeric(xFollowPeriod) Then
                    rsAdd("ES_RENEW_FOLLOWUP") = xFollowPeriod
                    'rsAdd("ES_FLWUP_PRD_DWMY") = "Y"
                    rsAdd("ES_FLWUP_PRD_DWMY") = UCase(Left(xFDWMY, 1))
                End If
            End If
            
            'rsAdd("ES_RENEW_CRS_CUR") = 99
            If Not IsEmpty(xRenewPeriod) Then
                If IsNumeric(xRenewPeriod) Then
                    rsAdd("ES_RENEW_CRS_CUR") = xRenewPeriod
                    'rsAdd("ES_CUR_PRD_DWMY") = "Y"
                    rsAdd("ES_CUR_PRD_DWMY") = UCase(Left(xCDWMY, 1))
                End If
            End If
            
            If UCase(Left(xUnique, 1)) = "Y" Then
                rsAdd("ES_UNIQUE_FOR_POS") = 1
            Else
                rsAdd("ES_UNIQUE_FOR_POS") = 0
            End If
            
            rsAdd("ES_LDATE") = Date
            rsAdd("ES_LTIME") = Time$
            rsAdd("ES_LUSER") = glbUserID
            rsAdd.Update
            
            xUptAmt = xUptAmt + 1
        
            'SQLQ = "ES_CRSCODE = " & rsCRSCodes("TB_KEY")
            'Ticket #24865 Franks 01/06/2013
            SQLQ = "ES_CRSCODE = '" & rsCRSCodes("TB_KEY") & "' "
            Data1.Recordset.Requery
            Data1.Recordset.Find SQLQ
        
            'Update employee's Training Plan
            Call Display_Value
            Call UpConttotal
            
            If xFollowPeriod = 99 And UCase(xFDWMY) = "YEAR" Then
                'Ticket #24865 Franks 01/06/2013
                'skip Add_Training_List_Rec_for_New_Renewal_Period
            Else
                Call Add_Training_List_Rec_for_New_Renewal_Period(rsCRSCodes("TB_KEY"))
            End If
            DoEvents
        End If
        rsCRSCodes.Close

NextRec:
    Next
    
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    Screen.MousePointer = vbDefault
    
    If xUptAmt > 0 Then spShow.FloodPercent = 100
    spShow.Visible = False
    
    If xUptAmt = 0 Then MsgBox "No record imported."
    If xUptAmt = 1 Then MsgBox "One record imported."
    If xUptAmt > 1 Then MsgBox Str(xUptAmt) & " records imported."
    
    Unload Me
    
End Sub

Private Function getRowsNspace(exSheet As Object, xFirstRow As Integer, Optional xCol, Optional NumSpa)
Dim X, K, m
X = xFirstRow
If IsMissing(xCol) Then
    K = 1
Else
    K = xCol
End If
If IsMissing(NumSpa) Then
    NumSpa = 1
End If
Do While True
    If exSheet.Cells(X, K) = "" Then
        m = m + 1
        If m > NumSpa Then
            Exit Do
        Else
            X = X + 1
        End If
    Else
        m = 1
        X = X + 1
    End If
Loop
getRowsNspace = X - 1
End Function


Private Sub cmdMissing_Click()
Dim SQLQ As String, Msg As String, a%
Dim rsCRSCodes As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim Inum As Integer

Msg = "This function will check what Course Codes are missing"
Msg = Msg & Chr(10) & "in Course Code Master, then populate these missing codes. "
Msg = Msg & Chr(10) & "Are you sure you want to do this?"
a% = MsgBox(Msg, 36, "Confirm")
If a% <> 6 Then Exit Sub

SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' "
SQLQ = SQLQ & "AND NOT (TB_KEY IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER)) "
rsCRSCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsCRSCodes.EOF Then
    MsgBox "No code added."
Else
    Inum = rsCRSCodes.RecordCount
    Do While Not rsCRSCodes.EOF
        SQLQ = "SELECT * FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & rsCRSCodes("TB_KEY") & "' "
        If rsAdd.State <> 0 Then rsAdd.Close
        rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsAdd.EOF Then
            rsAdd.AddNew
            rsAdd("ES_COMPNO") = "001"
            rsAdd("ES_CRSCODE") = rsCRSCodes("TB_KEY")
            If Not IsNull(rsCRSCodes("TB_USR1")) Then
                rsAdd("ES_CTYPE") = rsCRSCodes("TB_USR1")
            End If
            rsAdd("ES_STATUS") = 1
            rsAdd("ES_CORPONLY") = 0
            
            'Ticket #20840
            rsAdd("ES_UNIQUE_FOR_POS") = 0
            rsAdd("ES_RENEW_FOLLOWUP") = 99
            rsAdd("ES_FLWUP_PRD_DWMY") = "Y"

            rsAdd("ES_LDATE") = Date
            rsAdd("ES_LTIME") = Time$
            rsAdd("ES_LUSER") = glbUserID
            rsAdd.Update
        End If
        rsCRSCodes.MoveNext
    Loop
    MsgBox Str(Inum) & " codes added."
    Unload Me
End If
rsCRSCodes.Close
End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)

'Friesens - Ticket 16189
If glbCompSerial = "S/N - 2279W" And chkUnqforPos Then
    medCurPosRenewal.Text = ""
    medPrvPosRenewal.Text = ""
    medFlwUpEffective.Text = ""
    cmbCurDWMY.ListIndex = -1
    cmbPrvDWMY.ListIndex = -1
    cmbFlwUpDWMY.ListIndex = -1
    
    medCurPosRenewal.Enabled = False
    medPrvPosRenewal.Enabled = False
    medFlwUpEffective.Enabled = False
    cmbCurDWMY.Enabled = False
    cmbPrvDWMY.Enabled = False
    cmbFlwUpDWMY.Enabled = False
'7.9 - Enhancement - For all clients now
'ElseIf glbCompSerial = "S/N - 2188W" And chkUnqforPos Then  'City of Chatham-Kent - Ticket #16794
ElseIf chkUnqforPos Then  'City of Chatham-Kent - Ticket #16794
    medCurPosRenewal.Text = ""
    medFlwUpEffective.Text = ""
    cmbCurDWMY.ListIndex = -1
    cmbFlwUpDWMY.ListIndex = -1
    
    medCurPosRenewal.Enabled = False
    medFlwUpEffective.Enabled = False
    cmbCurDWMY.Enabled = False
    cmbFlwUpDWMY.Enabled = False
End If

'Friesens - Ticket 16189
If glbCompSerial = "S/N - 2279W" Then
    oCurRen = medCurPosRenewal.Text
    oPrvRen = medPrvPosRenewal.Text
    oFolRen = medFlwUpEffective.Text
    oCurRenTyp = txtCurDWMY.Text
    oPrvRenTyp = txtPrvDWMY.Text
    oFolRenTyp = txtFlwuUpDWMY.Text
'7.9 - Enhancement - For all clients now
Else 'If glbCompSerial = "S/N - 2188W" Then   'City of Chatham-Kent - Ticket #16794
    oCurRen = medCurPosRenewal.Text
    oFolRen = medFlwUpEffective.Text
    oCurRenTyp = txtCurDWMY.Text
    oFolRenTyp = txtFlwuUpDWMY.Text
End If

clpCode(0).SetFocus
fglbNewRec% = False

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

Private Sub cmdNew_Click()
Dim X%
glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True

Call Set_Control("B", Me)
rsDATA.AddNew

For X% = 0 To 4
    medEECont(X%) = 0
Next

chkStatus.Value = 1
chkCorponly = 0

Call UpConttotal

If glbWFC Then 'Ticket #21330 Franks 02/13/2012
    medFlwUpEffective.Text = 99
    txtFlwuUpDWMY.Text = "Y"
End If

clpCode(0).SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_COURSECODE_MASTER", "AddNew")
Resume Next

End Sub

Private Sub cmdOK_Click()
    Dim CourseCode
    Dim xRes As Integer
    
    On Error GoTo OK_Err
    
    'Data1.Refresh
    If Not chkCrsCode() Then Exit Sub
        
    Call UpdUStats(Me)
    
    CourseCode = clpCode(0)
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        txtCurDWMY.Text = Left(cmbCurDWMY, 1)
        txtPrvDWMY.Text = Left(cmbPrvDWMY, 1)
        txtFlwuUpDWMY.Text = Left(cmbFlwUpDWMY, 1)
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then   'City of Chatham-Kent - Ticket #16794
        txtCurDWMY.Text = Left(cmbCurDWMY, 1)
        txtFlwuUpDWMY.Text = Left(cmbFlwUpDWMY, 1)
    End If
    
    'Friesens - Ticket #16189
    If fglbNewRec% = False And glbCompSerial = "S/N - 2279W" And chkUnqforPos.Value = False Then
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oPrvRen <> medPrvPosRenewal.Text Or oPrvRenTyp <> txtPrvDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
        
            xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training List screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            If xRes = vbNo Then GoTo Skip_Save
        End If
    End If
        
    '7.9 - Enhancement - For all clients now
    'City of Chatham-Kent - Ticket #16794
    'If fglbNewRec% = False And glbCompSerial = "S/N - 2188W" And chkUnqforPos.Value = False Then
    If fglbNewRec% = False And glbCompSerial <> "S/N - 2279W" And chkUnqforPos.Value = False Then
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
        
            'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
            'Chatham-Kent but Chatham-Kent they are not using 7.9
            If glbCompSerial = "S/N - 2188W" Then
                xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training List screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            Else
                xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training Plan screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            End If
            
            If xRes = vbNo Then GoTo Skip_Save
        End If
    End If
    
    Screen.MousePointer = HOURGLASS
    
    Call Set_Control("U", Me, rsDATA)
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    
    'Friesens - Ticket #16189
    If fglbNewRec% = False And glbCompSerial = "S/N - 2279W" And chkUnqforPos.Value = False Then
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oPrvRen <> medPrvPosRenewal.Text Or oPrvRenTyp <> txtPrvDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
    
            'Check if renewal periods are being added instead of being modified
            If (oCurRen = "") And (medCurPosRenewal.Text <> "") Then
                'Current Renewal Period added
                Call Add_Training_List_Rec_for_New_Renewal_Period(clpCode(0).Text)
                
                'Check if Previous Renewal has been added as well
                If (oPrvRen = "") And (medPrvPosRenewal.Text <> "") Then
                    'Previous Renewal Period added
                    Call Add_Training_List_Rec_for_New_Prv_Renewal_Period(clpCode(0).Text)
                End If
            ElseIf (oPrvRen = "") And (medPrvPosRenewal.Text <> "") Then
                'Previous Renewal Period added
                Call Add_Training_List_Rec_for_New_Prv_Renewal_Period(clpCode(0).Text)
            End If
            
            'Changing from one value to another
            If (oCurRen <> "") Or (oPrvRen <> "") Then
                'Course renewal Period has changed - update Training List, Follow Up and Continuing Education records
                Call Update_Course_with_Changed_Renewal_Periods(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, medPrvPosRenewal.Text, txtPrvDWMY.Text, medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
            End If
            
            'Update Follow Up Renewal Periods
            If oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
                Call Update_Job_Course_Renewal(clpCode(0).Text)
            End If
        End If
    End If

    '7.9 - Enhancement - For all clients now
    'City of Chatham-Kent - Ticket #16794
    'If fglbNewRec% = False And glbCompSerial = "S/N - 2188W" And chkUnqforPos.Value = False Then
    If fglbNewRec% = False And glbCompSerial <> "S/N - 2279W" And chkUnqforPos.Value = False Then
        'Has the renewal periods changed?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
    
            'Check if renewal periods are being added instead of being modified
            If (oCurRen = "") And (medCurPosRenewal.Text <> "") Then
                'Current Renewal Period added
                Call Add_Training_List_Rec_for_New_Renewal_Period(clpCode(0).Text)
            End If
            
            'Changing from one value to another
            If (oCurRen <> "") Then
                'Course renewal Period has changed - update Training List, Follow Up and Continuing Education records
                Call Update_Course_with_Changed_Renewal_Periods(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, "", "", medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
            End If
            
            'Update Follow Up Renewal Periods
            If oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
                'Ticket #20518 - If Follow Up renewal is 99years then skip adding to training list
                If medFlwUpEffective.Text = "99" And txtFlwuUpDWMY.Text = "Y" Then
                    'Do nothing
                Else
                    Call Update_Job_Course_Renewal(clpCode(0).Text)
                End If
            End If
        End If
    End If


Skip_Save:
    Data1.Refresh
    Data1.Recordset.Find "ES_CRSCODE='" & CourseCode & " '"
    
    fglbNewRec% = False
    
    Call modSTUPD(False)
    
    Screen.MousePointer = DEFAULT
    
Exit Sub
    
OK_Err:
    Screen.MousePointer = DEFAULT
        
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_COURSECODE_MASTER", "Update")
    Resume Next
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

End Sub

Private Sub cmdSelect_Click()
glbCrsCode = clpCode(0).Text
glbCrsCodeDesc = clpCode(0).Caption
Call SaveDataInArray
Unload Me
End Sub

Private Sub Form_Activate()
glbOnTop = "frmMCourseCode"
End Sub

'Dim Ctrl As Control
Private Sub Form_Load()
    Dim SQLQ, i, ctylist, X
    
    glbOnTop = "frmMCourseCode"
    
    'Me.vbxTrueGrid.Columns(1).DataField = "ES_test"
    
    Data1.ConnectionString = glbAdoIHRDB
    'SQLQ = "SELECT * from HR_COURSECODE_MASTER WHERE (1=1) "
    'If glbCourseCodeSele Then
    '    SQLQ = SQLQ & "AND NOT (ES_STATUS = 0) "
    '    If Not glbDeptAllRight Then
    '        SQLQ = SQLQ & "AND (ES_CORPONLY = 0) "
    '    End If
    'End If
    'SQLQ = SQLQ & "ORDER BY ES_CRSCODE "
    
    SQLQ = "SELECT HR_COURSECODE_MASTER.*,HRTABL.TB_DESC AS COURSEDESC FROM HR_COURSECODE_MASTER, HRTABL WHERE HR_COURSECODE_MASTER.ES_CRSCODE_TABL = HRTABL.TB_NAME "
    SQLQ = SQLQ & "AND HRTABL.TB_NAME = 'ESCD' AND HR_COURSECODE_MASTER.ES_CRSCODE = HRTABL.TB_KEY "
    If glbCourseCodeSele Then
        SQLQ = SQLQ & "AND NOT (HR_COURSECODE_MASTER.ES_STATUS = 0) "
        If Not glbDeptAllRight Then
            SQLQ = SQLQ & "AND (HR_COURSECODE_MASTER.ES_CORPONLY = 0) "
        End If
    End If
    SQLQ = SQLQ & "ORDER BY ES_CRSCODE "

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Screen.MousePointer = HOURGLASS
    Me.vbxTrueGrid.Refresh
    Screen.MousePointer = DEFAULT
    
    Call modSTUPD(False)
    
    If Not gSec_Upd_CourseCodeMaster Then
        cmdModify.Enabled = False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
        cmdMissing.Enabled = False
    End If
    
    For i = 0 To 10
        Call setCaption(lbltitle(i))
    Next
    
    'Ticket #24708 Franks 11/27/2013
    Call setCaption(lbltitle(16))
    Call setCaption(lblCEUType)
    clpCEUType.TABLTitle = lStr("CEU Type")
    Call setCaption(lbltitle(24))
    
    For i = 0 To 9 '8
        vbxTrueGrid.Columns(i).Caption = lStr((vbxTrueGrid.Columns(i).Caption))
    Next i
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        lbltitle(12).Visible = True
        lbltitle(13).Visible = True
        lbltitle(14).Visible = True
        chkUnqforPos.Visible = True
        medCurPosRenewal.Visible = True
        medPrvPosRenewal.Visible = True
        medFlwUpEffective.Visible = True
        cmbCurDWMY.Visible = True
        cmbPrvDWMY.Visible = True
        cmbFlwUpDWMY.Visible = True
        
        If chkUnqforPos Then
            'lbltitle(12).FontBold = False
            'lbltitle(13).FontBold = False
            lbltitle(14).FontBold = False
        Else
            'lbltitle(12).FontBold = True
            'lbltitle(13).FontBold = True
            lbltitle(14).FontBold = True
        End If
    End If
    
    '7.9 - Enhancement - For all clients now
    'City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2188W" Then
    If glbCompSerial <> "S/N - 2279W" Then
        lbltitle(12).Visible = True
        lbltitle(13).Visible = False
        lbltitle(14).Visible = True
        chkUnqforPos.Visible = True
        medCurPosRenewal.Visible = True
        medPrvPosRenewal.Visible = False
        medFlwUpEffective.Visible = True
        cmbCurDWMY.Visible = True
        cmbPrvDWMY.Visible = False
        cmbFlwUpDWMY.Visible = True
        
        If chkUnqforPos Then
            lbltitle(14).FontBold = False
        Else
            lbltitle(14).FontBold = True
        End If
    End If
    
    If glbWFC Then 'Ticket #21330
        clpCode(3).DataField = "ES_WPSCODE"
        lbltitle(15).Top = lbltitle(13).Top
        clpCode(3).Top = medPrvPosRenewal.Top
        lbltitle(15).Visible = True
        clpCode(3).Visible = True
        'cmdImport.Visible = True 'Ticket #21851 Franks 04/19/2012, make Import generic
        
        'Ticket #24708 Franks 11/27/2013 - begin
        'lbltitle(16).Visible = True
        'clpCode(5).Visible = True
        'lblCEUType.Visible = True
        'clpCEUType.Visible = True
        'clpCEUType.DataField = "ES_CEUTYPE"
        'lbltitle(24).Visible = True
        'clpCode(4).Visible = True
        
        'lbltitle(16).FontBold = True 'Conducted By
        lbltitle(2).FontBold = True 'Coordinated By
        lbltitle(1).FontBold = True 'Course Type
        'lblCEUType.FontBold = True 'CEU Type
        lbltitle(24).FontBold = True 'Method Used
        imgHelpWFC.Visible = True
        lblShowHSECode.Visible = True
        'Ticket #24708 Franks 11/27/2013 - end
    End If
    
    'Call Display_Value
    
    Call INI_Controls(Me)
    
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        
        'Ticket #18210 Begin
        'If glbtermopen Then
        '    rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        'Else
        '    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'End If
        SQLQ = "SELECT * from HR_COURSECODE_MASTER WHERE (1=1) "
        If glbCourseCodeSele Then
            SQLQ = SQLQ & "AND NOT (ES_STATUS = 0) "
            If Not glbDeptAllRight Then
                SQLQ = SQLQ & "AND (ES_CORPONLY = 0) "
            End If
        End If
        SQLQ = SQLQ & "ORDER BY ES_CRSCODE "
        If glbtermopen Then
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        'Ticket #18210 - End
        'Call SET_UP_MODE
        Exit Sub
    End If
  
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER "
    SQLQ = SQLQ & " WHERE ES_CRSCODE='" & Data1.Recordset!ES_CRSCODE & "'" & " "
    If Not IsNull(Data1.Recordset!ES_CTYPE) Then
        SQLQ = SQLQ & " AND ES_CTYPE = '" & Data1.Recordset!ES_CTYPE & "'" & " "
    Else
        SQLQ = SQLQ & " AND ES_CTYPE IS NULL"
    End If
    SQLQ = SQLQ & " ORDER BY ES_CRSCODE"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)

    'glbCrsCode = clpCode(0).Text
    'glbCrsCodeDesc = clpCode(0).Caption
    
End Sub

Private Sub imgHelpWFC_Click()
Dim MsgStr As String
Dim xlsFileMat
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "HSE Code Help List.xls"
    If Dir(xlsFileMat) = "" Then
        MsgBox "There is no " & xlsFileMat
        Exit Sub
    End If
    
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
End Sub

Private Sub medEECont_LostFocus(Index As Integer)
Call UpConttotal
End Sub

Private Sub txtCourseHRS_LostFocus()
If Not IsNumeric(txtCourseHRS) Then txtCourseHRS = 0
If glbWFC Then 'Ticket #15522
    If fglbNewRec% Then
        If glbUNION = "NONE" Or glbUNION = "EXEC" Then
            medEECont(1).Text = txtCourseHRS * 50
        Else
            medEECont(1).Text = txtCourseHRS * 35
        End If
    End If
End If
End Sub

Private Sub txtCurDWMY_Change()
    cmbCurDWMY.ListIndex = -1
    Select Case txtCurDWMY
    Case "D"
        cmbCurDWMY.ListIndex = 0
    Case "W"
        cmbCurDWMY.ListIndex = 1
    Case "M"
        cmbCurDWMY.ListIndex = 2
    Case "Y"
        cmbCurDWMY.ListIndex = 3
    End Select
End Sub

Private Sub txtFindKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFlwuUpDWMY_Change()
    cmbFlwUpDWMY.ListIndex = -1
    Select Case txtFlwuUpDWMY
    Case "D"
        cmbFlwUpDWMY.ListIndex = 0
    Case "W"
        cmbFlwUpDWMY.ListIndex = 1
    Case "M"
        cmbFlwUpDWMY.ListIndex = 2
    Case "Y"
        cmbFlwUpDWMY.ListIndex = 3
    End Select
End Sub

Private Sub txtPrvDWMY_Change()
    cmbPrvDWMY.ListIndex = -1
    Select Case txtPrvDWMY
    Case "D"
        cmbPrvDWMY.ListIndex = 0
    Case "W"
        cmbPrvDWMY.ListIndex = 1
    Case "M"
        cmbPrvDWMY.ListIndex = 2
    Case "Y"
        cmbPrvDWMY.ListIndex = 3
    End Select
End Sub

Private Sub vbxTrueGrid_DblClick()
If Not Me.vbxTrueGrid.EditActive Then
    glbCrsCode = clpCode(0).Text
    glbCrsCodeDesc = clpCode(0).Caption
    Call SaveDataInArray
    Unload Me
Else
    MsgBox "Save/cancel changes first"
End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        

        SQLQ = "SELECT HR_COURSECODE_MASTER.*,HRTABL.TB_DESC AS COURSEDESC FROM HR_COURSECODE_MASTER, HRTABL WHERE HR_COURSECODE_MASTER.ES_CRSCODE_TABL = HRTABL.TB_NAME "
        SQLQ = SQLQ & "AND HRTABL.TB_NAME = 'ESCD' AND HR_COURSECODE_MASTER.ES_CRSCODE = HRTABL.TB_KEY "
        If glbCourseCodeSele Then
            SQLQ = SQLQ & "AND NOT (ES_STATUS = 0) "
            If Not glbDeptAllRight Then
                SQLQ = SQLQ & "AND (ES_CORPONLY = 0) "
            End If
        End If

        SQLQ = SQLQ & " ORDER BY  " & UCase(vbxTrueGrid.Columns(ColIndex).DataField) & " " & vbxTrueGrid.Tag

        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
    Call UpConttotal
    
    '7.9 - Enhancement - For all clients now
    'Friesens (Ticket #16189) and City of Chatham-Kent (Ticket #16794)
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        medCurPosRenewal.Enabled = False
        medPrvPosRenewal.Enabled = False
        medFlwUpEffective.Enabled = False
        cmbCurDWMY.Enabled = False
        cmbPrvDWMY.Enabled = False
        cmbFlwUpDWMY.Enabled = False
    'End If
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

cmdModify.Enabled = FT
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
vbxTrueGrid.Enabled = FT

txtCompanyName.Enabled = TF
txtTrainerName.Enabled = TF
txtCourseHRS.Enabled = TF
txtCompanyName.Enabled = TF
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
medEECont(0).Enabled = TF
medEECont(1).Enabled = TF
medEECont(2).Enabled = TF
medEECont(3).Enabled = TF
medEECont(4).Enabled = TF
medContTotal.Enabled = TF
clpEmpCur.Enabled = TF
clpOherCur.Enabled = TF
clpEmployerCur.Enabled = TF
clpAcomCur.Enabled = TF
clpLearnCur.Enabled = TF
clpLearnCur.Enabled = TF
clpTotCur.Enabled = TF
chkStatus.Enabled = TF
chkCorponly.Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF 'Ticket #24708 Franks 11/27/2013
clpCEUType.Enabled = TF 'Ticket #24708 Franks 11/27/2013
txtCEUCred.Enabled = TF 'Ticket #30365 Franks 07/26/2017

cmdClose.Enabled = FT           '
'cmdSelect.Enabled = FT          '
cmdPrint.Enabled = FT           '
cmdFind.Enabled = FT
txtFindKey.Enabled = FT
txtFindDesc.Enabled = FT

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False       '
End If

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    chkUnqforPos.Enabled = TF
    medCurPosRenewal.Enabled = TF
    medPrvPosRenewal.Enabled = TF
    medFlwUpEffective.Enabled = TF
    cmbCurDWMY.Enabled = TF
    cmbPrvDWMY.Enabled = TF
    cmbFlwUpDWMY.Enabled = TF
'7.9 - Enhancement - For all clients now
Else 'If glbCompSerial = "S/N - 2188W" Then   'City of Chatham-Kent - Ticket #16794
    chkUnqforPos.Enabled = TF
    medCurPosRenewal.Enabled = TF
    medPrvPosRenewal.Enabled = False
    medFlwUpEffective.Enabled = TF
    cmbCurDWMY.Enabled = TF
    cmbPrvDWMY.Enabled = False
    cmbFlwUpDWMY.Enabled = TF
End If
    
End Sub


Private Function chkCrsCode()
Dim CrsCode As String, SQLQ As String, Msg$
Dim snapCrsCodes As New ADODB.Recordset
Dim X
chkCrsCode = False
On Error GoTo chkCrsCode_Err

If Len(clpCode(0)) < 1 Then
    MsgBox lStr("Course Code is a required field")
    clpCode(0).SetFocus
    Exit Function
End If

If fglbNewRec% Then
    CrsCode = CStr(clpCode(0))
    SQLQ = "SELECT ES_CRSCODE, ES_ID from HR_COURSECODE_MASTER "
    SQLQ = SQLQ & "WHERE ES_CRSCODE = '" & CrsCode & "'"
    'SQLQ = SQLQ & " AND ES_ID = " & Data1.Recordset("ES_ID")
    
    If snapCrsCodes.State <> 0 Then snapCrsCodes.Close
    snapCrsCodes.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapCrsCodes.BOF And snapCrsCodes.EOF Then
        snapCrsCodes.Close
    Else
        'Ticket #20840 - When adding a new Course Code Master record, if a new Course Code is also added in the
        'lookup, then at that time a new Course Code Master for that Course is also added through IHRCTRLS.ocx.
        'This was asked by Jerry if the client is using Course Code Master. So to avoid confusion, we are
        'displaying the following message and also selecting that course code master record in the grid.

        SQLQ = "ES_ID = " & snapCrsCodes("ES_ID")
        Data1.Recordset.Requery
        Data1.Recordset.Find SQLQ
    
        Msg$ = lStr("This Course Code has already been added in the Course Code Master.")
        MsgBox Msg$
        
        Call cmdCancel_Click
        
        SQLQ = "ES_ID = " & snapCrsCodes("ES_ID")
        Data1.Recordset.Requery
        Data1.Recordset.Find SQLQ
        
        snapCrsCodes.Close
        
        Exit Function
    End If
End If

For X = 0 To 2
    If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X).SetFocus
        Exit Function
    End If
Next X

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" And chkUnqforPos.Value = False Then
    'If Len(Trim(medCurPosRenewal.Text)) = 0 Then
    '    MsgBox "Current Position Renewal Period cannot be blank"
    '    medCurPosRenewal.SetFocus
    '    Exit Function
    'End If
    If Len(Trim(medCurPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medCurPosRenewal.Text) Then
            MsgBox "Current Position Renewal Period is not numeric"
            medCurPosRenewal.SetFocus
            Exit Function
        End If
        If cmbCurDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Current Position Renewal Period"
            cmbCurDWMY.SetFocus
            Exit Function
        End If
    Else
        cmbCurDWMY.ListIndex = -1
    End If
    'If Len(Trim(medPrvPosRenewal.Text)) = 0 Then
    '    MsgBox "Previous Position Renewal Period cannot be blank"
    '    medPrvPosRenewal.SetFocus
    '    Exit Function
    'End If
    If Len(Trim(medPrvPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medPrvPosRenewal.Text) Then
            MsgBox "Previous Position Renewal Period is not numeric"
            medPrvPosRenewal.SetFocus
            Exit Function
        End If
        If cmbPrvDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Previous Position Renewal Period"
            cmbPrvDWMY.SetFocus
            Exit Function
        End If
    Else
         cmbPrvDWMY.ListIndex = -1
    End If
    
    If Len(Trim(medFlwUpEffective.Text)) = 0 Then
        MsgBox "Follow Up Effective Date Period cannot be blank"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If Not IsNumeric(medFlwUpEffective.Text) Then
        MsgBox "Follow Up Effective Date Period is not numeric"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If cmbFlwUpDWMY.ListIndex = -1 Then
        MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Follow Up Effective Date Period"
        cmbFlwUpDWMY.SetFocus
        Exit Function
    End If
ElseIf glbCompSerial = "S/N - 2279W" And chkUnqforPos Then
    medCurPosRenewal.Text = ""
    medPrvPosRenewal.Text = ""
    medFlwUpEffective.Text = ""
    cmbCurDWMY.ListIndex = -1
    cmbPrvDWMY.ListIndex = -1
    cmbFlwUpDWMY.ListIndex = -1
End If

'7.9 - Enhancement - For all clients now
'City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2188W" And chkUnqforPos.Value = False Then
If glbCompSerial <> "S/N - 2279W" And chkUnqforPos.Value = False Then
    If Len(Trim(medCurPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medCurPosRenewal.Text) Then
            MsgBox "Current Position Renewal Period is not numeric"
            medCurPosRenewal.SetFocus
            Exit Function
        End If
        If cmbCurDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Current Position Renewal Period"
            cmbCurDWMY.SetFocus
            Exit Function
        End If
    Else
        cmbCurDWMY.ListIndex = -1
    End If
    
    If Len(Trim(medFlwUpEffective.Text)) = 0 Then
        MsgBox "Follow Up Effective Date Period cannot be blank"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If Not IsNumeric(medFlwUpEffective.Text) Then
        MsgBox "Follow Up Effective Date Period is not numeric"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If cmbFlwUpDWMY.ListIndex = -1 Then
        MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Follow Up Effective Date Period"
        cmbFlwUpDWMY.SetFocus
        Exit Function
    End If
'7.9 - Enhancement - For all clients now
'ElseIf glbCompSerial = "S/N - 2188W" And chkUnqforPos Then
ElseIf glbCompSerial <> "S/N - 2279W" And chkUnqforPos Then
    medCurPosRenewal.Text = ""
    medFlwUpEffective.Text = ""
    cmbCurDWMY.ListIndex = -1
    cmbFlwUpDWMY.ListIndex = -1
End If

If glbWFC Then 'Ticket #24708 Franks 11/27/2013
    If Len(clpCode(1).Text) < 1 Then
        MsgBox lStr("Course Type") & " is a required field"
        clpCode(1).SetFocus
        Exit Function
    End If
    ''Ticket #24767 Franks 12/11/2013
    If Len(clpCode(2).Text) < 1 Then
        MsgBox lStr("Co-Ordinated By") & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
    'If Len(clpCode(5).Text) < 1 Then
    '    MsgBox lStr("Conducted By") & " is a required field"
    '    clpCode(5).SetFocus
    '    Exit Function
    'End If
    
    ''Ticket #24868 Franks 01/07/2013
    'If Len(clpCEUType.Text) < 1 Then
    '    MsgBox lStr("CEU Type") & " is a required field"
    '    clpCEUType.SetFocus
    '    Exit Function
    'End If
    If Len(clpCode(4).Text) < 1 Then
        MsgBox lStr("Method Used") & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
End If

chkCrsCode = True

Exit Function

chkCrsCode_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkCrsCode", "HR_COURSECODE_MASTER", "Cancel")
Resume Next

End Function

Private Sub UpConttotal()
Dim X%, xTotal
xTotal = ""
For X% = 0 To 4 '3
    If IsNumeric(medEECont(X%)) Then xTotal = Val(xTotal) + Val(medEECont(X%))
Next
medContTotal = xTotal

End Sub

Private Sub SaveDataInArray()
    glbCrsCodeStrArr(1) = clpCode(1).Text 'Course Type
    glbCrsCodeStrArr(2) = clpCode(2).Text 'Co-Ordinated By
    glbCrsCodeStrArr(3) = txtCompanyName.Text '
    glbCrsCodeStrArr(4) = txtTrainerName.Text '
    glbCrsCodeStrArr(5) = txtCourseHRS.Text 'Course Hours
    glbCrsCodeStrArr(6) = medEECont(0).Text 'Employee $
    glbCrsCodeStrArr(7) = medEECont(2).Text 'Other Expenses $
    glbCrsCodeStrArr(8) = medEECont(1).Text 'Employer $
    glbCrsCodeStrArr(9) = medEECont(3).Text 'Accommodation $
    glbCrsCodeStrArr(10) = medEECont(4).Text 'Learning Material $
    glbCrsCodeStrArr(11) = clpEmpCur.Text 'Currency
    glbCrsCodeStrArr(12) = clpOherCur.Text 'Currency
    glbCrsCodeStrArr(13) = clpEmployerCur.Text 'Currency
    glbCrsCodeStrArr(14) = clpAcomCur.Text 'Currency
    glbCrsCodeStrArr(15) = clpLearnCur.Text 'Currency
    glbCrsCodeStrArr(16) = clpTotCur.Text 'Currency
    glbCrsCodeStrArr(17) = "*" 'Flag
    'Ticket #24708 Franks 11/27/2013
    'glbCrsCodeStrArr(18) = clpCode(5).Text 'Conducted By
    glbCrsCodeStrArr(18) = clpCode(2).Text 'Coordinated By
    glbCrsCodeStrArr(19) = clpCEUType.Text  'CEU Type
    glbCrsCodeStrArr(20) = clpCode(4).Text 'Method Used

End Sub

Private Sub Deleted_Training_List_Records()
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    
    'If a course is deleted then the corresponding Position Required Courses and Training List records with this course
    'should be deleted as well. The follow up record should be marked Completed if the Course has been taken.
    'The Course Renewal Date on the Continuing Education screen should be cleared.
    'If TRAIN course then delete the Follow Up ref in HR_JOB_HISTORY and HR_TEMP_WORK tables.
    
    
    'Retrieve all training list records with this course
    SQLQ = "SELECT * FROM HR_TRAIN "
    SQLQ = SQLQ & " WHERE TR_CRSCODE = '" & clpCode(0).Text & "'"
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        'Records found in Training List with this Course
        rsHRTrain.MoveFirst
        
        Do While Not rsHRTrain.EOF
            'Clear the Renewal date for this course and for this employee from Continuing Education screen
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
            'if this is an independant from Position course in the Training List then Position Code should not
            'be part of the WHERE statement.
            If Not IsNull(rsHRTrain("TR_JOB")) And rsHRTrain("TR_JOB") <> "" Then
                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
            End If
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "'"
            SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(rsHRTrain("TR_RENEW"))
            SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                rsContEdu("ES_RENEW") = Null
                rsContEdu("ES_LDATE") = Date
                rsContEdu("ES_LUSER") = glbUserID
                rsContEdu("ES_LTIME") = Time$
                rsContEdu.Update
                
                If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                    'If valid Follow Up ID
                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                        'Since the course was completed - mark the Follow Up as
                        'Completed instead of deleting it.
                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        gdbAdoIhr001.Execute SQLQ
                    End If
                Else
                    'If valid Follow Up ID
                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                        'Delete the Follow Up record for this training record
                        'as no Course completion record found
                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        gdbAdoIhr001.Execute SQLQ
                    
                        'Clear the Follow Up ID in the Job History or Temp/Cross Training Position record
                        'if the course code is TRAIN
                        If rsHRTrain("TR_CRSCODE") = "TRAIN" Then
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Search HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("TW_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    End If
                End If
            Else
                'If valid Follow Up ID
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Delete the Follow Up record for this training record
                    'as no Course completion record found
                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    gdbAdoIhr001.Execute SQLQ
                
                    'Clear the Follow Up Id in the Job History or Temp/Cross Training Position record
                    'if the course code is TRAIN
                    If rsHRTrain("TR_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                        
                        'Search HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("TW_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
            End If
            rsContEdu.Close
            Set rsContEdu = Nothing
            
            'Delete this Training List record as the course is deleted from this position
            'SQLQ = "DELETE FROM HR_TRAIN"
            'SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR")
            'SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "'"
            'gdbAdoIhr001.Execute SQLQ
            rsHRTrain.Delete
            
            rsHRTrain.MoveNext
        Loop
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing
    
    'Delete the Course record from the Position's Required Courses list.
    SQLQ = "DELETE FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & clpCode(0).Text & "'"
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Private Sub Update_Course_with_Changed_Renewal_Periods(xCourseCode, Optional xCurRen, Optional xCurRenTyp, Optional xPrvRen, Optional xPrvRenTyp, Optional xFolRen, Optional xFolRenTyp)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ, xDWMY As String
    Dim oRenewalDate As Date
    Dim flgRenewalPeriod As Boolean
    Dim xComments As String

    'When Renewal Periods for Non 'Unique for each Position' courses change then these
    'courses renewal period should change on Position Required Course. Hence, it should
    'recompute the Renewal Date and update Training List, Continuing Education and
    'Follow Up records.
    
    'Also check if this course has been added in Training List as independant of job
    'course, update the Renewal Date if Course once taken on Training List, Follow Up
    'and Continuing Education

    'Retrieve Required Courses records with this course.
    SQLQ = "SELECT * FROM HR_JOB_COURSE"
    SQLQ = SQLQ & " WHERE PC_CRSCODE = '" & xCourseCode & "'"
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'For each Position, check which employee's have it as Current or marked it
            'as Track for Course Renewal in HR_JOB_HISTORY and HR_TEMP_WORK
            SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND JH_JOB = '" & rsReqCourse("PC_JOB") & "'"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
            SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND TW_JOB = '" & rsReqCourse("PC_JOB") & "'"
            rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJob.EOF Then
                rsEmpJob.MoveFirst
                
                Do While Not rsEmpJob.EOF
                    flgRenewalPeriod = True     'Renewal Period found - Default
                    
                    'Retrieve Training List records with this Job and Course
                    SQLQ = "SELECT * FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
                    SQLQ = SQLQ & " AND TR_JOB = '" & rsReqCourse("PC_JOB") & "'"
    
                    If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                        SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                    ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                        SQLQ = SQLQ & " AND TR_POS_TYPE = 'T'"
                    ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                        SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                    End If
                    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsHRTrain.EOF Then
                    
                        'Training List record found
                        oRenewalDate = rsHRTrain("TR_RENEW")    'Keep the original Renewal Date
                        flgRenewalPeriod = True                 'Renewal Period found - Default
                        
                        'Course Taken?
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Course Not Taken - Renewal Date based on Follow Up Period
                            Select Case xFolRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsHRTrain("TR_SDATE")))
                        Else
                            'Course Taken - Renewal Date based on the Renewal Period
                            'Check what type of Position it is and see if Renewal Period found for that
                            If rsEmpJob("TW_CURRENT") And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") Then
                                'Primary/Temporary Current Position - See if Current Renewal Period found
                                If Not IsNull(xCurRen) And xCurRen <> "" Then
                                    'Current Renewal Period found
                                    'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                    Select Case xCurRenTyp
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                Else
                                    'No Current Renewal Period
                                    flgRenewalPeriod = False
                                    
                                    'Delete Training List record, and update Follow Up and Continuing Education records
                                    GoTo Delete_Training_Record
                                End If
                            ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                                 'Previous Position - See if Previous Renewal Period found
                                If Not IsNull(xPrvRen) And xPrvRen <> "" Then
                                    'Previous Renewal Period found
                                    'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                    Select Case xPrvRenTyp
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                Else
                                    'No Previous Renewal Period
                                    flgRenewalPeriod = False
                                    
                                    'Delete Training List record, and update Follow Up and Continuing Education records
                                    GoTo Delete_Training_Record
                                End If
                            End If
                        End If
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain("TR_LTIME") = Time$
                        
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(oRenewalDate)
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                        
                        rsHRTrain.Update
                        
                                                
                        'Update Continuing Education record
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                        
                        
                        'Update Follow Up record
                        If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                    
Delete_Training_Record:
                        If flgRenewalPeriod = False Then
                            'Retrieve Continuing Education record
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                            SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    gdbAdoIhr001.Execute SQLQ
                                   
                                    'Clear the Follow Up Id on the Position record
                                    'if the course code is TRAIN
                                    If xCourseCode = "TRAIN" Then
                                        'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        If rsEmpJob("JOBTYPE") = "C" Then
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        End If
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            If rsEmpJob("JOBTYPE") = "C" Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                            ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                            End If
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Clear the Follow Up ID in the Position record
                                'if the course code is TRAIN
                                If xCourseCode = "TRAIN" Then
                                    'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    If rsEmpJob("JOBTYPE") = "C" Then
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    End If
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        If rsEmpJob("JOBTYPE") = "C" Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                        ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                        End If
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                                
                            'Delete this Training List record as the course
                            rsHRTrain.Delete
                        End If
                    
                    End If
                    rsHRTrain.Close
                    Set rsHRTrain = Nothing
                
                    rsEmpJob.MoveNext
                Loop
            End If
            rsEmpJob.Close
            Set rsEmpJob = Nothing
            
                        
            'Update Renewal Periods for this Course in this record
            rsReqCourse("PC_RENEW_CRS_CUR") = IIf(IsNull(xCurRen) Or xCurRen = "", Null, xCurRen)
            rsReqCourse("PC_RENEW_CRS_PRV") = IIf(IsNull(xPrvRen) Or xPrvRen = "", Null, xPrvRen)
            rsReqCourse("PC_RENEW_FOLLOWUP") = IIf(IsNull(xFolRen) Or xFolRen = "", Null, xFolRen)
            rsReqCourse("PC_CUR_PRD_DWMY") = xCurRenTyp
            rsReqCourse("PC_PRV_PRD_DWMY") = xPrvRenTyp
            rsReqCourse("PC_FLWUP_PRD_DWMY") = xFolRenTyp
            rsReqCourse.Update
            
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing
    
    
    
    'Check if this course is added to the Training List as independant course
    'If so, update the Renewal Date in Training List and Follow Up record, and if
    'course taken then in Continuing Education screen.
    
    'Renewal Date for the Courses Taken will be computed based on Current Renewal Period.
    'But when the Course is added for the first time it will be entered by User.
    
    'Retrieve Training List records with this Course and Job is null or blank
    SQLQ = "SELECT * FROM HR_TRAIN"
    SQLQ = SQLQ & " WHERE TR_CRSCODE = '" & xCourseCode & "'"
    SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '') "
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        rsHRTrain.MoveFirst
        
        Do While Not rsHRTrain.EOF
        
            'Training List record found
            oRenewalDate = rsHRTrain("TR_RENEW")    'Keep the original Renewal Date
            flgRenewalPeriod = True                 'Renewal Period found - Default
            
            'Course Taken?
            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                'Do not change the Renewal Date as this has been entered by the user
            Else
                'See if Current Renewal Period found
                If Not IsNull(xCurRen) And xCurRen <> "" Then
                    'Current Renewal Period found
                    'Calculate Renewal Date based on the Current Renewal Period and Course Taken Date
                    Select Case xCurRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                Else
                    'No Current Renewal Period
                    flgRenewalPeriod = False
                    
                    'Delete Training List record, and update Follow Up and Continuing Education records
                    GoTo Delete_Training_Record_1
                End If
                
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(oRenewalDate)
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                                
                rsHRTrain.Update
                
                'Update Continuing Education record
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                SQLQ = SQLQ & " AND (ES_JOB IS NULL OR ES_JOB = '')"
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                    rsContEdu("ES_LDATE") = Date
                    rsContEdu("ES_LUSER") = glbUserID
                    rsContEdu("ES_LTIME") = Time$
                    rsContEdu.Update
                End If
                rsContEdu.Close
                Set rsContEdu = Nothing
                
                'Update Follow Up record
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            
Delete_Training_Record_1:
                If flgRenewalPeriod = False Then
                    'Retrieve Continuing Education record
                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                    SQLQ = SQLQ & " AND (ES_JOB IS NULL OR ES_JOB = '')"
                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        
                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsContEdu.EOF Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete the Follow Up record for this training record
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            gdbAdoIhr001.Execute SQLQ
                        End If
                    Else
                        'Delete the Follow Up record for this training record
                        'as no Course record found
                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        gdbAdoIhr001.Execute SQLQ
                    End If
                    rsContEdu.Close
                    Set rsContEdu = Nothing
                        
                    'Delete this Training List record as the course
                    rsHRTrain.Delete
                End If  'if renewal period not found
                
            End If  'if course taken
            
            rsHRTrain.MoveNext
        Loop
    End If  'if training list record found
    rsHRTrain.Close
    Set rsHRTrain = Nothing
    
End Sub

Private Sub Add_Training_List_Rec_for_New_Renewal_Period(xCourseCode)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ, xDWMY, oJob As String
    Dim oRenewalDate As Date
    Dim flgChanged, flgCrsTakenBefore As Boolean
    Dim xComments As String

    'Renewal Period added to this course which was not existing before. Retrieve all the Jobs requiring this
    'course from the Required Courses table and then check which employee has this Job as Current or Tracked.
    'Job list should be ordered as Current, Temporary and Previous (Start Date Descending)
    'For all those jobs, check in the Training List based on the Type of Job - Current/Temp/Previous matching
    'the type of Renewal Period just added, if a Training List exists.
    'if CURRENT RENEWAL PERIOD added:
        'If the Course Taken is Blank then:
        
        '- If Type of Position is Current and the employee Position is Current - Skip to next record
        '- If Type of Position is Current and the employee Position is Temporary - Skip to next record
        '- If Type of Position is Current and the employee Position is Previous - Skip to next record
        'This is because Current Position takes precedence and Follow Up period has been used.
        
        '- If Type of Position is Temporary and the employee Position is Current
            '- change the Training List record Job and Type of Position to this Current Job. And renewal Period
            'based on the Current Job Position Start Date and Follow Up Period
        '- If Type of Position is Temporary and the employee Position is Temporary - Skip to next record
        '- If Type of Position is Temporary and the employee Position is Previous - Skip to next record
        
    
    'Retrieve Required Courses records with this course.
    SQLQ = "SELECT * FROM HR_JOB_COURSE"
    SQLQ = SQLQ & " WHERE PC_CRSCODE = '" & xCourseCode & "'"
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Retrieve employees with Job marked as Current only as Current Renewal has changed
            SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
            'SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE (JH_CURRENT <> 0)"
            SQLQ = SQLQ & " AND JH_JOB = '" & rsReqCourse("PC_JOB") & "'"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
            'SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE (TW_CURRENT <> 0)"
            SQLQ = SQLQ & " AND TW_JOB = '" & rsReqCourse("PC_JOB") & "'"
            SQLQ = SQLQ & " ORDER BY TW_EMPNBR, JOBTYPE ASC"
            rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJob.EOF Then
                rsEmpJob.MoveFirst
                
                Do While Not rsEmpJob.EOF
                    'Check in the Training List if this course exists
                    SQLQ = "SELECT * FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
                    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsHRTrain.EOF Then
                        'Retain the original date
                        oRenewalDate = rsHRTrain("TR_RENEW")
                        oJob = rsHRTrain("TR_JOB")
                        flgChanged = False
                        
                        'For PRIMACY or TEMPORARY Current type of Jobs
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Course Not Taken
                            If rsHRTrain("TR_POS_TYPE") <> "C" Then
                                'Course had not been taken and it's not a Current Type Training List record,
                                'reset the Renewal Date
                                Select Case UCase(Left(cmbFlwUpDWMY, 1))
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If Not IsNull(rsEmpJob("TW_SDATE")) Then 'Ticket #24074 Franks 07/16/2013
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medFlwUpEffective.Text, CVDate(rsEmpJob("TW_SDATE")))
                                End If
                                flgChanged = True
                            End If
                        Else
                            'Course Taken
                            If rsHRTrain("TR_POS_TYPE") <> "C" Then
                                'Course Taken by another Job Type - Current takes the precedence
                                'Recompute the Renewal Date for Current Job
                                If medCurPosRenewal.Text <> "" Then
                                    Select Case UCase(Left(cmbCurDWMY, 1))
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medCurPosRenewal.Text, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    flgChanged = True
                                Else
                                    flgChanged = False
                                End If
                            End If
                        End If
                        
                        If flgChanged = True Then
                            'Change took place - update the rest of the fields and table
                            rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                            rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                rsHRTrain("TR_POS_TYPE") = "C"
                            ElseIf (rsEmpJob("JOBTYPE") = "T") Then
                                rsHRTrain("TR_POS_TYPE") = "T"
                            End If
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(oRenewalDate)
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            rsHRTrain.Update
                            
                            'Update Continuing Education record
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                'if Course Taken
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                    rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            End If
                            
                            'Update Follow Up record
                            If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If xCourseCode = "TRAIN" Then
                                'Clear the Follow Up ID from the older job record
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                Else
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If (rsEmpJob("JOBTYPE") = "C") Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                    Else
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                                Else
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If (rsEmpJob("JOBTYPE") = "C") Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    Else
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    Else
                        'No Training List records found for this Job
                        'Add Training List record
                        flgCrsTakenBefore = False
                        
                        rsHRTrain.AddNew
                        rsHRTrain("TR_COMPNO") = "001"
                        rsHRTrain("TR_EMPNBR") = rsEmpJob("TW_EMPNBR")
                        rsHRTrain("TR_CRSCODE") = xCourseCode
                        
                        'Check first if this Course was taken before in the Continuing Education screen
                        flgCrsTakenBefore = False
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                        SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            'Course Taken Before
                            rsContEdu.MoveFirst
                            flgCrsTakenBefore = True
                        Else
                            'Course not taken before
                            flgCrsTakenBefore = False
                        
                            'Ticket #19816
                            'Search for Cont Edu with Renewal Date
                            '7.9 - Enhancement - For all clients now
                            'If glbCompSerial = "S/N - 2188W" Then
                            If glbCompSerial <> "S/N - 2279W" Then
                                'Renewal Date is not null
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                SQLQ = SQLQ & " AND (ES_RENEW IS NOT NULL)"
                                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'Course Taken Before
                                    rsContEdu.MoveFirst
                                    flgCrsTakenBefore = True
                                Else
                                    'Course not taken before
                                    flgCrsTakenBefore = False
                                End If
                            End If
                        End If
                        
                        If flgCrsTakenBefore = True Then
                            'Course Taken Before
                            'Compute the Renewal Date based on last Course Taken Date and Current Renewal Period
                            If medCurPosRenewal.Text <> "" Then
                                Select Case UCase(Left(cmbCurDWMY, 1))
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                
                                '7.9 - Enhancement - For all clients now
                                'Ticket #19816
                                'If glbCompSerial = "S/N - 2188W" Then
                                If glbCompSerial <> "S/N - 2279W" Then
                                    If IsDate(rsContEdu("ES_RENEW")) Then
                                        rsHRTrain("TR_RENEW") = rsContEdu("ES_RENEW")   'If they have already entered the Renewal Date then follow that.
                                    Else
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medCurPosRenewal.Text, CVDate(rsContEdu("ES_DATCOMP")))
                                    End If
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medCurPosRenewal.Text, CVDate(rsContEdu("ES_DATCOMP")))
                                End If
                                rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                                
                                'Update Continuing Education record as well
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                        Else
                            'Course Not Taken
                            'Compute Renewal Date based on Follow Up Renewal Period
                            Select Case UCase(Left(cmbFlwUpDWMY, 1))
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If Not IsNull(rsEmpJob("TW_SDATE")) Then 'Ticket #24074 Franks 07/16/2013
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medFlwUpEffective.Text, CVDate(rsEmpJob("TW_SDATE")))
                            End If
                        End If
                        
                        rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                        rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf (rsEmpJob("JOBTYPE") = "T") Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        End If
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LTIME") = Time$
                        rsHRTrain("TR_LUSER") = glbUserID
                        
                        'Add a Follow Up record for this Training course
                        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        rsFollowUp.AddNew
                        rsFollowUp("EF_COMPNO") = "001"
                        rsFollowUp("EF_EMPNBR") = rsEmpJob("TW_EMPNBR")
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        rsFollowUp("EF_FREAS_TABL") = "FURE"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                            rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJob("TW_EMPNBR"), "ED_ADMINBY", Null)
                        End If
                        rsFollowUp("EF_FREAS") = "EDUC"
                        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp.Update
                        
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                        rsHRTrain.Update
                        
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                    
                        'Update Position record with Follow Up ID
                        'if the course code is TRAIN
                        If xCourseCode = "TRAIN" Then
                            'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                            Else
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                            End If
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                Else
                                    rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                    rsHRTrain.Close
                    Set rsHRTrain = Nothing
                    
                    rsEmpJob.MoveNext
                Loop
            End If
            rsEmpJob.Close
            Set rsEmpJob = Nothing
            
            'Update Renewal Periods for this Course in this record
            rsReqCourse("PC_RENEW_CRS_CUR") = IIf(IsNull(medCurPosRenewal.Text) Or medCurPosRenewal.Text = "", Null, medCurPosRenewal.Text)
            rsReqCourse("PC_RENEW_CRS_PRV") = IIf(IsNull(medPrvPosRenewal.Text) Or medPrvPosRenewal.Text = "", Null, medPrvPosRenewal.Text)
            rsReqCourse("PC_RENEW_FOLLOWUP") = IIf(IsNull(medFlwUpEffective.Text) Or medFlwUpEffective.Text = "", Null, medFlwUpEffective.Text)
            rsReqCourse("PC_CUR_PRD_DWMY") = UCase(Left(cmbCurDWMY, 1))
            rsReqCourse("PC_PRV_PRD_DWMY") = UCase(Left(cmbPrvDWMY, 1))
            rsReqCourse("PC_FLWUP_PRD_DWMY") = UCase(Left(cmbFlwUpDWMY, 1))
            rsReqCourse.Update
            
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing
End Sub

Private Sub Add_Training_List_Rec_for_New_Prv_Renewal_Period(xCourseCode)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ, xDWMY, oJob As String
    Dim oRenewalDate, lstEndDate As Date
    Dim flgChanged, flgCrsTakenBefore As Boolean
    Dim lstEmpNo As Integer
    Dim xComments As String
    
    'Renewal Period added to this course which was not existing before. Retrieve all the Jobs requiring this
    'course from the Required Courses table and then check which employee has this Job as Current or Tracked.
    'Job list should be ordered as Current, Temporary and Previous (Start Date Descending)
    'For all those jobs, check in the Training List based on the Type of Job - Current/Temp/Previous matching
    'the type of Renewal Period just added, if a Training List exists.
    'if PREVIOUS RENEWAL PERIOD added:
        'If the Course Taken is Blank then:
                
    
    'Retrieve Required Courses records with this course.
    SQLQ = "SELECT * FROM HR_JOB_COURSE"
    SQLQ = SQLQ & " WHERE PC_CRSCODE = '" & xCourseCode & "'"
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Retrieve employees with Job marked as Current only as Current Renewal has changed
            SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
            'SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE (JH_TRK_CRS_RENEWAL <> 0)"
            SQLQ = SQLQ & " AND JH_JOB = '" & rsReqCourse("PC_JOB") & "'"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
            'SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE (TW_TRK_CRS_RENEWAL <> 0)"
            SQLQ = SQLQ & " AND TW_JOB = '" & rsReqCourse("PC_JOB") & "'"
            SQLQ = SQLQ & " ORDER BY TW_EMPNBR, TW_ENDDATE DESC"
            rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJob.EOF Then
                rsEmpJob.MoveFirst
                
                Do While Not rsEmpJob.EOF
                    'Check in the Training List if this course exists
                    SQLQ = "SELECT * FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
                    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsHRTrain.EOF Then
                        'Retain the original date
                        oRenewalDate = rsHRTrain("TR_RENEW")
                        oJob = rsHRTrain("TR_JOB")
                        flgChanged = False
                        
                        'Last record
                        If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                            lstEmpNo = rsEmpJob("TW_EMPNBR")
                            lstEndDate = rsEmpJob("TW_ENDDATE")
                        End If
                        
                        'For Previous type of Jobs
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Course Not Taken
                            If rsHRTrain("TR_POS_TYPE") <> "C" And rsHRTrain("TR_POS_TYPE") <> "T" Then
                                If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                                    'Course had not been taken and it's not a Current/Temp Type Training List record,
                                    'reset the Renewal Date
                                    Select Case UCase(Left(cmbFlwUpDWMY, 1))
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    If Not IsNull(rsEmpJob("TW_SDATE")) Then 'Ticket #24074 Franks 07/16/2013
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medFlwUpEffective.Text, CVDate(rsEmpJob("TW_SDATE")))
                                    End If
                                    flgChanged = True
                                Else
                                    flgChanged = False
                                End If
                            End If
                        Else
                            'Course Taken
                            If rsHRTrain("TR_POS_TYPE") <> "C" And rsHRTrain("TR_POS_TYPE") <> "T" Then
                                If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                                    'Course Taken by another Job Type - Current/Temp takes the precedence
                                    'Recompute the Renewal Date for Previous Job
                                    If medPrvPosRenewal.Text <> "" Then
                                        Select Case UCase(Left(cmbPrvDWMY, 1))
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medPrvPosRenewal.Text, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                        flgChanged = True
                                    Else
                                        flgChanged = False
                                    End If
                                Else
                                    flgChanged = False
                                End If
                            End If
                        End If
                        
                        If flgChanged = True Then
                            'Change took place - update the rest of the fields and table
                            rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                            rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                            rsHRTrain("TR_POS_TYPE") = "P"
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(oRenewalDate)
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            rsHRTrain.Update
                            
                            'Update Continuing Education record
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                'if Course Taken
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                    rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            End If
                            
                            'Update Follow Up record
                            If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If xCourseCode = "TRAIN" Then
                                'Clear the Follow Up ID from the older job record
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                Else
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If (rsEmpJob("JOBTYPE") = "C") Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                    Else
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                                Else
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If (rsEmpJob("JOBTYPE") = "C") Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    Else
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    Else
                        'No Training List records found for this Job
                        'Add Training List record
                        flgCrsTakenBefore = False
                        
                        rsHRTrain.AddNew
                        rsHRTrain("TR_COMPNO") = "001"
                        rsHRTrain("TR_EMPNBR") = rsEmpJob("TW_EMPNBR")
                        rsHRTrain("TR_CRSCODE") = xCourseCode
                        
                        'Check first if this Course was taken before in the Continuing Education screen
                        flgCrsTakenBefore = False
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                        SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            'Course Taken Before
                            rsContEdu.MoveFirst
                            flgCrsTakenBefore = True
                        Else
                            'Course not taken before
                            flgCrsTakenBefore = False
                        End If
                        
                        If flgCrsTakenBefore = True Then
                            'Course Taken Before
                            'Compute the Renewal Date based on last Course Taken Date and Current Renewal Period
                            If medPrvPosRenewal.Text <> "" Then
                                Select Case UCase(Left(cmbPrvDWMY, 1))
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medPrvPosRenewal.Text, CVDate(rsContEdu("ES_DATCOMP")))
                                rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                                
                                'Update Continuing Education record as well
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                        Else
                            'Course Not Taken
                            'Compute Renewal Date based on Follow Up Renewal Period
                            Select Case UCase(Left(cmbFlwUpDWMY, 1))
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If Not IsNull(rsEmpJob("TW_SDATE")) Then 'Ticket #24074 Franks 07/16/2013
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, medFlwUpEffective.Text, CVDate(rsEmpJob("TW_SDATE")))
                            End If
                        End If
                        
                        rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                        rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                        rsHRTrain("TR_POS_TYPE") = "P"
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LTIME") = Time$
                        rsHRTrain("TR_LUSER") = glbUserID
                        
                        'Add a Follow Up record for this Training course
                        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        rsFollowUp.AddNew
                        rsFollowUp("EF_COMPNO") = "001"
                        rsFollowUp("EF_EMPNBR") = rsEmpJob("TW_EMPNBR")
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        rsFollowUp("EF_FREAS_TABL") = "FURE"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                            rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJob("TW_EMPNBR"), "ED_ADMINBY", Null)
                        End If
                        rsFollowUp("EF_FREAS") = "EDUC"
                        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp.Update
                        
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                        rsHRTrain.Update
                        
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                    
                        'Update Position record with Follow Up ID
                        'if the course code is TRAIN
                        If xCourseCode = "TRAIN" Then
                            'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                            Else
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                            End If
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                If (rsEmpJob("JOBTYPE") = "C") Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                Else
                                    rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                    rsHRTrain.Close
                    Set rsHRTrain = Nothing
                    
                    rsEmpJob.MoveNext
                Loop
            End If
            rsEmpJob.Close
            Set rsEmpJob = Nothing
            
            'Update Renewal Periods for this Course in this record
            rsReqCourse("PC_RENEW_CRS_CUR") = IIf(IsNull(medCurPosRenewal.Text) Or medCurPosRenewal.Text = "", Null, medCurPosRenewal.Text)
            rsReqCourse("PC_RENEW_CRS_PRV") = IIf(IsNull(medPrvPosRenewal.Text) Or medPrvPosRenewal.Text = "", Null, medPrvPosRenewal.Text)
            rsReqCourse("PC_RENEW_FOLLOWUP") = IIf(IsNull(medFlwUpEffective.Text) Or medFlwUpEffective.Text = "", Null, medFlwUpEffective.Text)
            rsReqCourse("PC_CUR_PRD_DWMY") = UCase(Left(cmbCurDWMY, 1))
            rsReqCourse("PC_PRV_PRD_DWMY") = UCase(Left(cmbPrvDWMY, 1))
            rsReqCourse("PC_FLWUP_PRD_DWMY") = UCase(Left(cmbFlwUpDWMY, 1))
            rsReqCourse.Update
            
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing
End Sub

Private Sub Update_Job_Course_Renewal(xCourseCode)
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String
    
    'Retrieve Required Courses records with this course.
    SQLQ = "SELECT * FROM HR_JOB_COURSE"
    SQLQ = SQLQ & " WHERE PC_CRSCODE = '" & xCourseCode & "'"
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Update Renewal Periods for this Course in this record
            rsReqCourse("PC_RENEW_CRS_CUR") = IIf(IsNull(medCurPosRenewal.Text) Or medCurPosRenewal.Text = "", Null, medCurPosRenewal.Text)
            rsReqCourse("PC_RENEW_CRS_PRV") = IIf(IsNull(medPrvPosRenewal.Text) Or medPrvPosRenewal.Text = "", Null, medPrvPosRenewal.Text)
            rsReqCourse("PC_RENEW_FOLLOWUP") = IIf(IsNull(medFlwUpEffective.Text) Or medFlwUpEffective.Text = "", Null, medFlwUpEffective.Text)
            rsReqCourse("PC_CUR_PRD_DWMY") = UCase(Left(cmbCurDWMY, 1))
            rsReqCourse("PC_PRV_PRD_DWMY") = UCase(Left(cmbPrvDWMY, 1))
            rsReqCourse("PC_FLWUP_PRD_DWMY") = UCase(Left(cmbFlwUpDWMY, 1))
            rsReqCourse.Update
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

End Sub

Private Sub CourseCode_Type()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String, RType
Dim rsTABL As New ADODB.Recordset
'''On Error GoTo Dept_GL_Err

If Len(clpCode(0).Text) > 0 Then
    rsTABL.Open "SELECT TB_NAME,TB_KEY,TB_USR1 FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY='" & clpCode(0).Text & "'", gdbAdoIhr001
    If Not rsTABL.EOF Then
        If IsNull(rsTABL("TB_USR1")) Then
            RType = ""
        Else
            RType = rsTABL("TB_USR1")
        End If
        If Len(RType) > 0 Then
            If glbWFC Then 'Ticket #25676 Franks 07/08/2014
                '"   Since Woodbridge always has the Course Type, don't display this message.
            Else
                If clpCode(1).Text <> RType Then
                    Msg$ = lStr("Do you want the associated Course Type?")
                    Title$ = "info:HR"
                    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2   ' Describe dialog.
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDYES Then clpCode(1).Text = RType
                End If
            End If
        End If
    End If
    rsTABL.Close
    Set rsTABL = Nothing
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
    ' clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Course Code Snap", "Course Code Master", "SELECT")
Call RollBack '21June99 js
End Sub

'Sub CrsName_Desc()
'    'If Course Code is blank, don't wipe up the Course Name
'    If Len(clpCode(0).Caption) > 0 Then
'        'Frank 10/20/03
'        'As Jerry request, if Course Name exists there, don't replace it
'        If Len(Trim(txtCourseName)) = 0 Then
'            txtCourseName = Replace(clpCode(0).Caption, "&&", "&")
'        End If
'    End If
'End Sub
