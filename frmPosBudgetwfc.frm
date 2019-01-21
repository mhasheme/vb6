VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmPosBudgetWFC 
   Caption         =   "Budgeted Positions"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   5535
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDelete 
      Height          =   1455
      Left            =   9960
      TabIndex        =   92
      Top             =   6600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdDeleYear 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   375
         Left            =   1440
         TabIndex        =   97
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtYearD 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   285
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   93
         Tag             =   "01-Number of positions that exist for this job"
         Top             =   240
         Width           =   855
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   94
         Top             =   600
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.Frame frmCopyPlant 
      Height          =   4335
      Left            =   10560
      TabIndex        =   79
      Top             =   2160
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdCopyPlant 
         Appearance      =   0  'Flat
         Caption         =   "Copy"
         Height          =   375
         Left            =   1320
         TabIndex        =   90
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox txtYearT 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   285
         Left            =   1995
         MaxLength       =   4
         TabIndex        =   73
         Tag             =   "01-Number of positions that exist for this job"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtYearF 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   285
         Left            =   1995
         MaxLength       =   4
         TabIndex        =   69
         Tag             =   "01-Number of positions that exist for this job"
         Top             =   480
         Width           =   855
      End
      Begin INFOHR_Controls.CodeLookup clpDivF 
         Height          =   285
         Left            =   1680
         TabIndex        =   71
         Tag             =   "00-Specific Division Desired"
         Top             =   1200
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpDeptF 
         Height          =   285
         Left            =   1680
         TabIndex        =   72
         Top             =   1560
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   70
         Top             =   840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpDivT 
         Height          =   285
         Left            =   1680
         TabIndex        =   75
         Tag             =   "00-Specific Division Desired"
         Top             =   3000
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpDeptT 
         Height          =   285
         Left            =   1680
         TabIndex        =   76
         Top             =   3360
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   74
         Top             =   2640
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   89
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   88
         Top             =   3405
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   87
         Top             =   3045
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   86
         Top             =   2280
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Copy To:"
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
         Left            =   120
         TabIndex        =   85
         Top             =   1960
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Copy From:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   84
         Top             =   160
         Width           =   960
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   83
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   82
         Top             =   1605
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   81
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   80
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.Frame frmCopyOne 
      Height          =   680
      Left            =   4080
      TabIndex        =   67
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdCopyOne 
         Appearance      =   0  'Flat
         Caption         =   "Copy"
         Height          =   375
         Left            =   3240
         TabIndex        =   78
         Top             =   180
         Width           =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   68
         Tag             =   "40-Status To Date"
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frmCopy 
      Height          =   1215
      Left            =   120
      TabIndex        =   62
      Top             =   6600
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton OptCopyRec 
         Caption         =   "Copy one plant/division/department/year  "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   66
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton OptCopyRec 
         Caption         =   "Copy one plant/division/year  "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton OptCopyRec 
         Caption         =   "Copy one plant/year "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton OptCopyRec 
         Caption         =   "Copy this record"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      DataField       =   "JG_YEAR"
      Height          =   285
      Left            =   2235
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "01-Number of positions that exist for this job"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox medFTEHrs 
      Appearance      =   0  'Flat
      DataField       =   "JG_FTEHRS"
      Height          =   285
      Left            =   2235
      TabIndex        =   11
      Tag             =   "10-FTE Hours/Year"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox medFTENum 
      Appearance      =   0  'Flat
      DataField       =   "JG_FTENUM"
      Height          =   285
      Left            =   2235
      TabIndex        =   10
      Tag             =   "10-Number of FTE "
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtNoPos 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      DataField       =   "JG_BUDGNBR"
      Height          =   285
      Left            =   2235
      MaxLength       =   3
      TabIndex        =   9
      Tag             =   "01-Number of positions that exist for this job"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   12120
      MaxLength       =   25
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   11070
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   9840
      MaxLength       =   25
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   14715
      _Version        =   65536
      _ExtentX        =   25956
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
      Begin VB.Label lblSecDesc 
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
         Left            =   7800
         TabIndex        =   52
         Top             =   135
         Width           =   2535
      End
      Begin VB.Label lblSec 
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
         Left            =   6720
         TabIndex        =   51
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   50
         Top             =   165
         Width           =   690
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
         Left            =   2640
         TabIndex        =   28
         Top             =   135
         Width           =   3015
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
         Left            =   960
         TabIndex        =   27
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   165
         Width           =   690
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   14715
      _Version        =   65536
      _ExtentX        =   25956
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
      Begin VB.CommandButton cmdCopy2NextYear 
         Appearance      =   0  'Flat
         Caption         =   "Copy to Next Year"
         Height          =   375
         Left            =   10080
         TabIndex        =   24
         Top             =   195
         Width           =   2055
      End
      Begin VB.CommandButton cmdCountOnePos 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate 1 Position"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   195
         Width           =   2265
      End
      Begin VB.CommandButton cmdExp 
         Appearance      =   0  'Flat
         Caption         =   "Export Into Excel File"
         Height          =   375
         Left            =   7920
         TabIndex        =   23
         Top             =   195
         Width           =   2055
      End
      Begin VB.CommandButton cmdImp 
         Appearance      =   0  'Flat
         Caption         =   "Import From Excel File"
         Height          =   375
         Left            =   5760
         TabIndex        =   22
         Top             =   195
         Width           =   2055
      End
      Begin VB.CommandButton cmdCountPos 
         Appearance      =   0  'Flat
         Caption         =   "Recalculate all Positions"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   195
         Width           =   2265
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   13560
         TabIndex        =   19
         Tag             =   "Print Listing "
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4425
         TabIndex        =   18
         Tag             =   "Delete the Record Selected"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3585
         TabIndex        =   17
         Tag             =   "Add a new Record"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2685
         TabIndex        =   16
         Tag             =   "Cancel the changes made"
         Top             =   195
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12600
         TabIndex        =   15
         Tag             =   "Save the changes made"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   14280
         TabIndex        =   14
         Tag             =   "Edit the information on this screen"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   12840
         TabIndex        =   13
         Tag             =   "Close and exit this screen"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   14760
         Top             =   480
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   13440
         Top             =   480
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
      Begin Crystal.CrystalReport vbxCrystal2 
         Left            =   14640
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
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPosBudgetwfc.frx":0000
      Height          =   2115
      Left            =   0
      OleObjectBlob   =   "frmPosBudgetwfc.frx":0014
      TabIndex        =   12
      Tag             =   "Skills Lookup"
      Top             =   600
      Width           =   11115
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "JG_DIV"
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Specific Division Desired"
      Top             =   3960
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "JG_DEPTNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   4320
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpGLNum 
      DataField       =   "JG_GLNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Tag             =   "00-General Ledger - Code"
      Top             =   4680
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin Threed.SSCheck chkCurrent 
      DataField       =   "JG_CURRENT"
      Height          =   255
      Left            =   4630
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1890
      _Version        =   65536
      _ExtentX        =   3334
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Current"
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
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JG_JREASON"
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Tag             =   "01-Reason code "
      Top             =   3600
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BPRC"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "JG_FRDATE"
      DataSource      =   " "
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Tag             =   "40-Status From Date"
      Top             =   3240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "JG_TODATE"
      DataSource      =   " "
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Tag             =   "40-Status To Date"
      Top             =   3240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "JG_EFDATE"
      DataSource      =   " "
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Tag             =   "40-Status To Date"
      Top             =   3600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin Threed.SSCheck chkCopy 
      Height          =   255
      Left            =   120
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6360
      Width           =   2730
      _Version        =   65536
      _ExtentX        =   4815
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Copy Budgeted Positions"
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
   End
   Begin Threed.SSCheck chkDeleYearPlant 
      Height          =   255
      Left            =   4080
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   6360
      Width           =   2730
      _Version        =   65536
      _ExtentX        =   4815
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Delete Year/Plant"
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
   End
   Begin VB.Image imgPosFilled 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3240
      Picture         =   "frmPosBudgetwfc.frx":890C
      Stretch         =   -1  'True
      Top             =   5070
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_VACANCY_POS"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6960
      TabIndex        =   60
      Top             =   5085
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   59
      Top             =   3285
      Width           =   195
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3600
      TabIndex        =   57
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacancy Positions"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   56
      Top             =   5085
      Width           =   1425
   End
   Begin VB.Label lblEDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date "
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
      TabIndex        =   55
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblDateRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
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
      TabIndex        =   54
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label lblYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      TabIndex        =   53
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label lblSections 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      DataField       =   "JG_SECTION"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11040
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   540
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   4005
      Width           =   1410
   End
   Begin VB.Label lblFTEHrs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTE Hours/Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   47
      Top             =   5805
      Width           =   1275
   End
   Begin VB.Label lblTotHrs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total FTE Hours/Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3600
      TabIndex        =   46
      Top             =   5805
      Width           =   1575
   End
   Begin VB.Label lblFTETotHrs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTETOTHR"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5190
      TabIndex        =   45
      Top             =   5805
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTENUMVACN"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6960
      TabIndex        =   44
      Top             =   5445
      Width           =   570
   End
   Begin VB.Label lblPosFiled 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_NBRFIL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      TabIndex        =   43
      Top             =   5085
      Width           =   570
   End
   Begin VB.Label lblFTETotNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTENUMFILL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      TabIndex        =   42
      Top             =   5445
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacancy # FTE"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   5520
      TabIndex        =   41
      Top             =   5445
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTE # Filled"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   40
      Top             =   5445
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTE #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   39
      Top             =   5445
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Positions Filled"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   38
      Top             =   5085
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Budgeted #Pos'ns"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   37
      Top             =   5085
      Width           =   1440
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11040
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblPositions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POST"
      DataField       =   "JG_CODE"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8160
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "JG_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8160
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   4365
      Width           =   1560
   End
   Begin VB.Label lblGLNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   4725
      Width           =   870
   End
End
Attribute VB_Name = "frmPosBudgetWFC"
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
Dim fglbFDate, fglbTDate, fglbEDate

Private Sub chkCopy_Click(Value As Integer)
    If chkCopy.Value Then
        If chkDeleYearPlant.Value Then
            chkDeleYearPlant.Value = False
        End If
        frmDelete.Visible = False
    End If
    If chkCopy.Value Then
        frmCopy.Visible = True
    Else
        frmCopy.Visible = False
        frmCopyPlant.Visible = False
        frmCopyOne.Visible = False
        OptCopyRec(0).Value = False
        OptCopyRec(1).Value = False
        OptCopyRec(2).Value = False
        OptCopyRec(3).Value = False
    End If
End Sub

Private Sub chkDeleYearPlant_Click(Value As Integer)
    If chkDeleYearPlant.Value Then
        If chkCopy.Value Then
            chkCopy.Value = False
        End If
        frmCopy.Visible = False
        frmCopyPlant.Visible = False
        frmDelete.Visible = True
    Else
        frmDelete.Visible = False
    End If
End Sub

Private Sub clpDept_Change()
    Call Dept_GL
End Sub

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


Private Sub cmdCopy2NextYear_Click() 'Ticket #28254 Franks 04/05/2016
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String
Dim xNum As Integer
Dim xRows As Long
Dim xRow As Long
Dim xEmpnbr
Dim xFlag As Boolean
Dim xYear, xNextYear, xPlant, xPos As String, xDIV, xBUnit, xDept As String, xGL As String, xBudPos, xFte, xFTEHrs, xUptMsg
Dim xTmp
    
    If Not gSec_Upd_BudgetedPos Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    xYear = txtYear.Text
    If Not IsNumeric(xYear) Then
        MsgBox "   Invalid Year "
        Exit Sub
    End If
    If Not Len(xYear) = 4 Then
        MsgBox "   Invalid Year "
        Exit Sub
    End If
    xNextYear = xYear + 1
    
    Msg = "This program will copy all current Budgeted Positions of " & xYear & " into " & xNextYear & " "
    Msg = Msg & Chr(10) & "also reset the Current flags for " & xNextYear & " only."
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do it? "
    a% = MsgBox(Msg, 36, "Confirm")
    
    If a% <> 6 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    If CopyRecordsAll(xYear, xNextYear) Then
        Data1.Refresh
        MsgBox "   Finished.   "
    End If
    Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdCopyOne_Click()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, INo&
    If Len(dlpDate(3).Text) < 1 Then
        MsgBox "Effective Date is a required field"
        dlpDate(3).SetFocus
        Exit Sub
    Else
        If Not IsDate(dlpDate(3).Text) Then
            MsgBox "Invalid Date"
            dlpDate(3).SetFocus
            Exit Sub
        End If
        If CVDate((dlpDate(3).Text)) < CVDate((dlpDate(0).Text)) Then
            MsgBox "Effective Date must be within the Date Range"
            dlpDate(3).SetFocus
            Exit Sub
        End If
        If CVDate((dlpDate(3).Text)) > CVDate((dlpDate(1).Text)) Then
            MsgBox "Effective Date must be within the Date Range"
            dlpDate(3).SetFocus
            Exit Sub
        End If
    End If
    If CVDate((dlpDate(3).Text)) = CVDate((dlpDate(2).Text)) Then
        MsgBox "New Effective Date can't be same as the current Effective Date"
        dlpDate(3).SetFocus
        Exit Sub
    End If
    
    Msg = "Are you sure you want to copy this record? "
    a% = MsgBox(Msg, 36, "Confirm Copy")
    
    If a% <> 6 Then Exit Sub

    SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & glbPos & "' "
    SQLQ = SQLQ & "AND JG_SECTION = '" & glbJobSection & "' "
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & "AND JG_DIV = '" & clpDiv.Text & "' "
    End If
    If Len(clpDept.Text) > 0 Then
        SQLQ = SQLQ & "AND JG_DEPTNO = '" & clpDept.Text & "' "
    End If
    If Len(clpGLNum.Text) > 0 Then
        SQLQ = SQLQ & "AND JG_GLNO = '" & clpGLNum.Text & "' "
    End If
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYear.Text & " "
    SQLQ = SQLQ & "AND JG_EFDATE = " & Date_SQL(dlpDate(3).Text) & " "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        rs("JG_COMPNO") = "001"
        rs("JG_CODE") = glbPos
        rs("JG_BUDPOSNBR") = rsDATA("JG_BUDPOSNBR")
        rs("JG_DIV") = rsDATA("JG_DIV")
        rs("JG_DEPTNO") = rsDATA("JG_DEPTNO")
        rs("JG_GLNO") = rsDATA("JG_GLNO")
        rs("JG_NBRFIL") = rsDATA("JG_NBRFIL")
        rs("JG_BUDGNBR") = rsDATA("JG_BUDGNBR")
        rs("JG_FTENUM") = rsDATA("JG_FTENUM")
        rs("JG_FTENUMFILL") = rsDATA("JG_FTENUMFILL")
        rs("JG_FTENUMVACN") = rsDATA("JG_FTENUMVACN")
        rs("JG_FTEHRS") = rsDATA("JG_FTEHRS")
        rs("JG_FTETOTHR") = rsDATA("JG_FTETOTHR")
        rs("JG_LDATE") = Date
        rs("JG_LTIME") = Time$
        rs("JG_LUSER") = glbUserID
        rs("JG_SECTION") = rsDATA("JG_SECTION")
        rs("JG_YEAR") = rsDATA("JG_YEAR")
        rs("JG_FRDATE") = rsDATA("JG_FRDATE")
        rs("JG_TODATE") = rsDATA("JG_TODATE")
        rs("JG_EFDATE") = dlpDate(3).Text
        rs("JG_JREASON") = rsDATA("JG_JREASON")
        rs("JG_VACANCY_POS") = rsDATA("JG_VACANCY_POS")
        rs.Update
    End If
    rs.Close
    Call setCurrentFlag(glbPos, glbJobSection)
    Data1.Refresh
End Sub

Private Sub cmdCopyPlant_Click()
    If OptCopyRec(1).Value Then
        Call CopyPlant1
    End If
    If OptCopyRec(2).Value Then
        Call CopyPlant2
    End If
    If OptCopyRec(3).Value Then
        Call CopyPlant3
    End If
End Sub
Private Sub CopyPlant3()
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, INo&
Dim xFlag As Boolean

    If Len(txtYearF.Text) < 1 Then
        MsgBox ("From Year is a required field")
        txtYearF.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearF.Text) Then
            MsgBox "From Year must be numeric."
            txtYearF.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearF.Text) = 4 Then
            MsgBox "Invalid From Year."
            txtYearF.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(1).Text) < 1 Then
        MsgBox "From " & lStr("Section") & (" is a required field")
        clpCode(1).SetFocus
        Exit Sub
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox "From " & lStr("Section") & (" must be valid")
            clpCode(1).SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDivF.Text) < 1 Then
        MsgBox "From " & lStr("Division") & (" is a required field")
        clpDivF.SetFocus
        Exit Sub
    Else
        If clpDivF.Caption = "Unassigned" Then
            MsgBox "From " & lStr("Division") & (" must be valid")
            clpDivF.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDeptF.Text) < 1 Then
        MsgBox "From " & lStr("Department") & (" is a required field")
        clpDeptF.SetFocus
        Exit Sub
    Else
        If clpDeptF.Caption = "Unassigned" Then
            MsgBox "From " & lStr("Department") & (" must be valid")
            clpDeptF.SetFocus
            Exit Sub
        End If
    End If
    
    'Copy To:
    If Len(txtYearT.Text) < 1 Then
        MsgBox ("To Year is a required field")
        txtYearT.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearT.Text) Then
            MsgBox "To Year must be numeric."
            txtYearT.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearT.Text) = 4 Then
            MsgBox "Invalid To Year."
            txtYearT.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "To " & lStr("Section") & (" is a required field")
        clpCode(2).SetFocus
        Exit Sub
    Else
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "To " & lStr("Section") & (" must be valid")
            clpCode(2).SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDivT.Text) < 1 Then
        MsgBox "To " & lStr("Division") & (" is a required field")
        clpDivT.SetFocus
        Exit Sub
    Else
        If clpDivT.Caption = "Unassigned" Then
            MsgBox "To " & lStr("Division") & (" must be valid")
            clpDivT.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDeptT.Text) < 1 Then
        MsgBox "To " & lStr("Department") & (" is a required field")
        clpDeptT.SetFocus
        Exit Sub
    Else
        If clpDeptT.Caption = "Unassigned" Then
            MsgBox "To " & lStr("Department") & (" must be valid")
            clpDeptT.SetFocus
            Exit Sub
        End If
    End If
    
    'check if there are record based on From fields
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearF.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(1).Text & "' "
    SQLQ = SQLQ & "AND JG_DIV = '" & clpDivF.Text & "' "
    SQLQ = SQLQ & "AND JG_DEPTNO = '" & clpDeptF.Text & "' "
    If rsFrom.State <> 0 Then rsFrom.Close
    rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsFrom.EOF Then
            MsgBox "No record found based on Copy From fields. "
            txtYearF.SetFocus
            rsFrom.Close
            Exit Sub
    End If
    rsFrom.Close
    
    xFlag = False
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearT.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(2).Text & "' "
    SQLQ = SQLQ & "AND JG_DIV = '" & clpDivT.Text & "' "
    SQLQ = SQLQ & "AND JG_DEPTNO = '" & clpDeptT.Text & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
            Msg = "Found records based on Copy To fields. "
            Msg = Msg & Chr(10) & "Are you sure you want to update these record? "
            a% = MsgBox(Msg, 36, "Confirm ")
            If a% <> 6 Then
                txtYearF.SetFocus
                Exit Sub
            Else
                xFlag = True
            End If
    End If
    rs.Close
    
    If Not xFlag Then
        Msg = "Are you sure you want to copy these record? "
        a% = MsgBox(Msg, 36, "Confirm Copy")
        If a% <> 6 Then Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    Call CopyRecords(txtYearF.Text, clpCode(1).Text, clpDivF.Text, clpDeptF.Text, txtYearT.Text, clpCode(2).Text, clpDivT.Text, clpDeptT.Text)
    Call CalcCurFlagForPlant(clpCode(2).Text)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    
    MsgBox "   Finished.   "

End Sub
Private Sub CopyPlant2()
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, INo&
Dim xFlag As Boolean

    If Len(txtYearF.Text) < 1 Then
        MsgBox ("From Year is a required field")
        txtYearF.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearF.Text) Then
            MsgBox "From Year must be numeric."
            txtYearF.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearF.Text) = 4 Then
            MsgBox "Invalid From Year."
            txtYearF.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(1).Text) < 1 Then
        MsgBox "From " & lStr("Section") & (" is a required field")
        clpCode(1).SetFocus
        Exit Sub
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox "From " & lStr("Section") & (" must be valid")
            clpCode(1).SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDivF.Text) < 1 Then
        MsgBox "From " & lStr("Division") & (" is a required field")
        clpDivF.SetFocus
        Exit Sub
    Else
        If clpDivF.Caption = "Unassigned" Then
            MsgBox "From " & lStr("Division") & (" must be valid")
            clpDivF.SetFocus
            Exit Sub
        End If
    End If
    
    'Copy To:
    If Len(txtYearT.Text) < 1 Then
        MsgBox ("To Year is a required field")
        txtYearT.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearT.Text) Then
            MsgBox "To Year must be numeric."
            txtYearT.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearT.Text) = 4 Then
            MsgBox "Invalid To Year."
            txtYearT.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "To " & lStr("Section") & (" is a required field")
        clpCode(2).SetFocus
        Exit Sub
    Else
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "To " & lStr("Section") & (" must be valid")
            clpCode(2).SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDivT.Text) < 1 Then
        MsgBox "To " & lStr("Division") & (" is a required field")
        clpDivT.SetFocus
        Exit Sub
    Else
        If clpDivT.Caption = "Unassigned" Then
            MsgBox "To " & lStr("Division") & (" must be valid")
            clpDivT.SetFocus
            Exit Sub
        End If
    End If
    
    
    'check if there are record based on From fields
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearF.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(1).Text & "' "
    SQLQ = SQLQ & "AND JG_DIV = '" & clpDivF.Text & "' "
    If rsFrom.State <> 0 Then rsFrom.Close
    rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsFrom.EOF Then
            MsgBox "No record found based on Copy From fields. "
            txtYearF.SetFocus
            rsFrom.Close
            Exit Sub
    End If
    rsFrom.Close
    
    xFlag = False
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearT.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(2).Text & "' "
    SQLQ = SQLQ & "AND JG_DIV = '" & clpDivT.Text & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
            Msg = "Found records based on Copy To fields. "
            Msg = Msg & Chr(10) & "Are you sure you want to update these record? "
            a% = MsgBox(Msg, 36, "Confirm ")
            If a% <> 6 Then
                txtYearF.SetFocus
                Exit Sub
            Else
                xFlag = True
            End If
    End If
    rs.Close
    
    If Not xFlag Then
        Msg = "Are you sure you want to copy these record? "
        a% = MsgBox(Msg, 36, "Confirm Copy")
        If a% <> 6 Then Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    Call CopyRecords(txtYearF.Text, clpCode(1).Text, clpDivF.Text, clpDeptF.Text, txtYearT.Text, clpCode(2).Text, clpDivT.Text, clpDeptT.Text)
    Call CalcCurFlagForPlant(clpCode(2).Text)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    
    MsgBox "   Finished.   "

End Sub
Private Sub CopyPlant1()
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, INo&
Dim xFlag As Boolean

    If Len(txtYearF.Text) < 1 Then
        MsgBox ("From Year is a required field")
        txtYearF.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearF.Text) Then
            MsgBox "From Year must be numeric."
            txtYearF.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearF.Text) = 4 Then
            MsgBox "Invalid From Year."
            txtYearF.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(1).Text) < 1 Then
        MsgBox "From " & lStr("Section") & (" is a required field")
        clpCode(1).SetFocus
        Exit Sub
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox "From " & lStr("Section") & (" must be valid")
            clpCode(1).SetFocus
            Exit Sub
        End If
    End If
    If Len(txtYearT.Text) < 1 Then
        MsgBox ("To Year is a required field")
        txtYearT.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearT.Text) Then
            MsgBox "To Year must be numeric."
            txtYearT.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearT.Text) = 4 Then
            MsgBox "Invalid To Year."
            txtYearT.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "To " & lStr("Section") & (" is a required field")
        clpCode(2).SetFocus
        Exit Sub
    Else
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "To " & lStr("Section") & (" must be valid")
            clpCode(2).SetFocus
            Exit Sub
        End If
    End If
    
    
    'check if there are record based on From fields
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearF.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(1).Text & "' "
    If rsFrom.State <> 0 Then rsFrom.Close
    rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsFrom.EOF Then
            MsgBox "No record found based on Copy From fields. "
            txtYearF.SetFocus
            rsFrom.Close
            Exit Sub
    End If
    rsFrom.Close
    
    xFlag = False
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearT.Text & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(2).Text & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
            Msg = "Found records based on Copy To fields. "
            Msg = Msg & Chr(10) & "Are you sure you want to update these record? "
            a% = MsgBox(Msg, 36, "Confirm ")
            If a% <> 6 Then
                txtYearF.SetFocus
                Exit Sub
            Else
                xFlag = True
            End If
    End If
    rs.Close
    
    If Not xFlag Then
        Msg = "Are you sure you want to copy these record? "
        a% = MsgBox(Msg, 36, "Confirm Copy")
        If a% <> 6 Then Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    Call CopyRecords(txtYearF.Text, clpCode(1).Text, clpDivF.Text, clpDeptF.Text, txtYearT.Text, clpCode(2).Text, clpDivT.Text, clpDeptT.Text)
    Call CalcCurFlagForPlant(clpCode(2).Text)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    
    MsgBox "   Finished.   "
End Sub
Private Function CopyRecordsAll(xYearF, xYearT)  ', xPlantF, xDivF, xDeptF, xYearT, xPlantT, xDivT, xDeptT)
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPlantT, xDivT, xDeptT
Dim xNum, xRows
    CopyRecordsAll = False
    
    'Ticket #29484 Franks 11/21/2016 -
    'There are some positions with 2017 already. The copy program should skip those records and just copy the old year to the new year. This message isn't required.
    'Thanks.  Jerry
    'check if xYearT exist
    ''SQLQ = "SELECT TOP 10 * FROM HRJOBBUD WHERE (1=1) "
    ''SQLQ = SQLQ & "AND JG_YEAR = " & xYearT & " "
    ''If rsFrom.State <> 0 Then rsFrom.Close
    ''rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    ''If Not rsFrom.EOF Then
    ''    MsgBox xYearT & " Budgeted Positions already exist. Can not create these records again. "
    ''    Exit Function
    ''End If
    ''rsFrom.Close
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    Call getFromToDate(xYearT) 'get the Copy To From, To and Eff Dates

    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & xYearF & " "
    SQLQ = SQLQ & "AND NOT JG_CURRENT = 0 "
    'SQLQ = SQLQ & "AND JG_SECTION = '" & xPlantF & "' "
    'If Len(xDivF) > 0 Then
    '    SQLQ = SQLQ & "AND JG_DIV = '" & xDivF & "' "
    'End If
    'If Len(xDeptF) > 0 Then
    '    SQLQ = SQLQ & "AND JG_DEPTNO = '" & xDeptF & "' "
    'End If
    SQLQ = SQLQ & "ORDER BY JG_SECTION,JG_DIV,JG_DEPTNO,JG_CODE,JG_YEAR,JG_EFDATE "
    If rsFrom.State <> 0 Then rsFrom.Close
    rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    xNum = 0
    If Not rsFrom.EOF Then
        xRows = rsFrom.RecordCount
    End If
    Do While Not rsFrom.EOF
        MDIMain.panHelp(0).FloodPercent = (xNum / xRows) * 100
        xNum = xNum + 1
        DoEvents
    
        'xPlantT, xDivT, xDeptT
        If Not IsNull(rsFrom("JG_SECTION")) Then xPlantT = rsFrom("JG_SECTION") Else xPlantT = ""
        If Not IsNull(rsFrom("JG_DIV")) Then xDivT = rsFrom("JG_DIV") Else xDivT = ""
        If Not IsNull(rsFrom("JG_DEPTNO")) Then xDeptT = rsFrom("JG_DEPTNO") Else xDeptT = ""
        
        SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & rsFrom("JG_CODE") & "' "
        SQLQ = SQLQ & "AND JG_YEAR = " & xYearT & " "
        SQLQ = SQLQ & "AND JG_SECTION = '" & xPlantT & "' "
        If Len(xDivT) > 0 Then
            SQLQ = SQLQ & "AND JG_DIV = '" & xDivT & "' "
        End If
        If Len(xDeptT) > 0 Then
            SQLQ = SQLQ & "AND JG_DEPTNO = '" & xDeptT & "' "
        End If
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            rs.AddNew
            rs("JG_COMPNO") = "001"
            rs("JG_CODE") = rsFrom("JG_CODE")  'glbPos
        End If
        rs("JG_BUDPOSNBR") = rsFrom("JG_BUDPOSNBR")
        If Len(xDivT) = 0 Then
            rs("JG_DIV") = rsFrom("JG_DIV")  'Null
        Else
            rs("JG_DIV") = xDivT '
        End If
        If Len(xDeptT) = 0 Then
            rs("JG_DEPTNO") = rsFrom("JG_DEPTNO") 'Null
        Else
            rs("JG_DEPTNO") = xDeptT '
        End If
        rs("JG_GLNO") = rsFrom("JG_GLNO")
        rs("JG_NBRFIL") = rsFrom("JG_NBRFIL")
        rs("JG_BUDGNBR") = rsFrom("JG_BUDGNBR")
        rs("JG_FTENUM") = rsFrom("JG_FTENUM")
        rs("JG_FTENUMFILL") = rsFrom("JG_FTENUMFILL")
        rs("JG_FTENUMVACN") = rsFrom("JG_FTENUMVACN")
        rs("JG_FTEHRS") = rsFrom("JG_FTEHRS")
        rs("JG_FTETOTHR") = rsFrom("JG_FTETOTHR")
        rs("JG_LDATE") = Date
        rs("JG_LTIME") = Time$
        rs("JG_LUSER") = glbUserID
        rs("JG_SECTION") = xPlantT 'rsFrom("JG_SECTION")
        rs("JG_YEAR") = xYearT 'rsFrom("JG_YEAR")
        rs("JG_FRDATE") = fglbFDate ' rsFrom("JG_FRDATE")
        rs("JG_TODATE") = fglbTDate 'rsFrom("JG_TODATE")
        rs("JG_EFDATE") = fglbEDate 'rsFrom("JG_EFDATE")
        rs("JG_JREASON") = "" 'rsFrom("JG_JREASON")
        rs("JG_VACANCY_POS") = rsFrom("JG_VACANCY_POS")
        rs("JG_CURRENT") = 1 ' rsFrom("JG_CURRENT")
        rs.Update
        rs.Close
    
        rsFrom.MoveNext
    Loop
    rsFrom.Close
    
    'reset the upload flag to false for Non xYearT year
    SQLQ = "UPDATE HRJOBBUD SET JG_CURRENT = 0 WHERE NOT (JG_YEAR = " & xYearT & ") "
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    CopyRecordsAll = True
End Function

Private Sub CopyRecords(xYearF, xPlantF, xDivF, xDeptF, xYearT, xPlantT, xDivT, xDeptT)
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String

    Call getFromToDate(xYearT) 'get the Copy To From, To and Eff Dates

    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & xYearF & " "
    SQLQ = SQLQ & "AND JG_SECTION = '" & xPlantF & "' "
    If Len(xDivF) > 0 Then
        SQLQ = SQLQ & "AND JG_DIV = '" & xDivF & "' "
    End If
    If Len(xDeptF) > 0 Then
        SQLQ = SQLQ & "AND JG_DEPTNO = '" & xDeptF & "' "
    End If
    SQLQ = SQLQ & "ORDER BY JG_SECTION,JG_DIV,JG_DEPTNO,JG_CODE,JG_YEAR,JG_EFDATE "
    If rsFrom.State <> 0 Then rsFrom.Close
    rsFrom.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsFrom.EOF
        SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & rsFrom("JG_CODE") & "' "
        SQLQ = SQLQ & "AND JG_YEAR = " & xYearT & " "
        SQLQ = SQLQ & "AND JG_SECTION = '" & xPlantT & "' "
        If Len(xDivT) > 0 Then
            SQLQ = SQLQ & "AND JG_DIV = '" & xDivT & "' "
        End If
        If Len(xDeptT) > 0 Then
            SQLQ = SQLQ & "AND JG_DEPTNO = '" & xDeptT & "' "
        End If
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            rs.AddNew
            rs("JG_COMPNO") = "001"
            rs("JG_CODE") = glbPos
        End If
        rs("JG_BUDPOSNBR") = rsFrom("JG_BUDPOSNBR")
        If Len(xDivT) = 0 Then
            rs("JG_DIV") = rsFrom("JG_DIV")  'Null
        Else
            rs("JG_DIV") = xDivT '
        End If
        If Len(xDeptT) = 0 Then
            rs("JG_DEPTNO") = rsFrom("JG_DEPTNO") 'Null
        Else
            rs("JG_DEPTNO") = xDeptT '
        End If
        rs("JG_GLNO") = rsFrom("JG_GLNO")
        rs("JG_NBRFIL") = rsFrom("JG_NBRFIL")
        rs("JG_BUDGNBR") = rsFrom("JG_BUDGNBR")
        rs("JG_FTENUM") = rsFrom("JG_FTENUM")
        rs("JG_FTENUMFILL") = rsFrom("JG_FTENUMFILL")
        rs("JG_FTENUMVACN") = rsFrom("JG_FTENUMVACN")
        rs("JG_FTEHRS") = rsFrom("JG_FTEHRS")
        rs("JG_FTETOTHR") = rsFrom("JG_FTETOTHR")
        rs("JG_LDATE") = Date
        rs("JG_LTIME") = Time$
        rs("JG_LUSER") = glbUserID
        rs("JG_SECTION") = xPlantT 'rsFrom("JG_SECTION")
        rs("JG_YEAR") = xYearT 'rsFrom("JG_YEAR")
        rs("JG_FRDATE") = fglbFDate ' rsFrom("JG_FRDATE")
        rs("JG_TODATE") = fglbTDate 'rsFrom("JG_TODATE")
        rs("JG_EFDATE") = fglbEDate 'rsFrom("JG_EFDATE")
        rs("JG_JREASON") = "" 'rsFrom("JG_JREASON")
        rs("JG_VACANCY_POS") = rsFrom("JG_VACANCY_POS")
        rs("JG_CURRENT") = rsFrom("JG_CURRENT")
        rs.Update
        rs.Close
    
        rsFrom.MoveNext
    Loop
    rsFrom.Close
    


End Sub

Private Sub cmdCountOnePos_Click()
Dim a As Integer, Msg As String, INo&

On Error GoTo CountErr

Msg = "Are You Sure You Want To Count Budgeted Position? "
a% = MsgBox(Msg, 36, "Confirm")

If a% <> 6 Then Exit Sub

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    Call setCurrentFlag(glbPos, glbJobSection)
    Call mod_Upd_Pos_Budget_WFC(glbPos, glbJobSection)
    Data1.Refresh
    Call Display_Value
    MsgBox "Budgeted Position Counted"
Else
    MsgBox "No Position Selected."
End If

Exit Sub

CountErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Count Pos Error", "HRJOBBUD Refresh", "Refresh")
Resume Next
Call RollBack
End Sub

Private Sub cmdCountPos_Click()
Dim a As Integer, Msg As String, INo&
On Error GoTo CountErr

Msg = "Are You Sure You Want To Count Budgeted Positions? "
a% = MsgBox(Msg, 36, "Confirm")

If a% <> 6 Then Exit Sub

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If mod_Upd_Pos_Budget_WFC("", "") Then
        Beep
        MsgBox "Budgeted Positions Counted"
    End If
    Data1.Refresh
    Call Display_Value
End If

Exit Sub

CountErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Count Pos Error", "HRJOBBUD Refresh", "Refresh")
Resume Next
Call RollBack

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&

If Not gSec_Upd_BudgetedPos Then
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
fglbNew = False
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Call setCurrentFlag(glbPos, glbJobSection)
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

Private Sub cmdDeleYear_Click()
Dim rsFrom As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, INo&
Dim xFlag As Boolean
Dim I As Integer
    If Len(txtYearD.Text) < 1 Then
        MsgBox ("Year is a required field")
        txtYearD.SetFocus
        Exit Sub
    Else
        If Not IsNumeric(txtYearD.Text) Then
            MsgBox "Year must be numeric."
            txtYearD.SetFocus
            Exit Sub
        End If
        If Not Len(txtYearD.Text) = 4 Then
            MsgBox "Invalid From Year."
            txtYearD.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(3).Text) < 1 Then
        'MsgBox lStr("Section") & (" is a required field")
        'clpCode(3).SetFocus
        'Exit Sub
    Else
        If clpCode(3).Caption = "Unassigned" Then
            MsgBox lStr("Section") & (" must be valid")
            clpCode(3).SetFocus
            Exit Sub
        End If
    End If
    
    Msg = "Are you sure you want to delete these record? "
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
        
    SQLQ = "DELETE FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_YEAR = " & txtYearD.Text & " "
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & "AND JG_SECTION = '" & clpCode(3).Text & "' "
    End If
    gdbAdoIhr001.Execute SQLQ, I
    
    If I = 0 Then
        MsgBox "No record deleted."
    Else
        If Len(clpCode(3).Text) = 0 Then
            Call CalcCurFlagForAll
        Else
            Call CalcCurFlagForPlant(clpCode(3).Text)
        End If
        MsgBox I & " record(s) deleted."
    End If
    

    
    Data1.Refresh

End Sub

Private Sub cmdExp_Click()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rs As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim ImportFile, xlsFileTmp
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String
Dim xNum As Integer
Dim xRows As Long
Dim xRow As Long
Dim xEmpnbr
Dim xFlag As Boolean
Dim xYear, xPlant, xPos As String, xDIV, xBUnit, xDept As String, xGL As String, xBudPos, xFte, xFTEHrs, xUptMsg
Dim xTmp

    If Not gSec_Upd_BudgetedPos Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    'Budgeted Position Import.xls or Budgeted Position Import.xlsx
    'check file name
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Budgeted Position Export Tmp.xls"
    ImportFile = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Budgeted Position Export.xls"
    If Dir(xlsFileTmp) = "" Then
      MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
      Exit Sub
    End If
    
    Msg = "This program will export Budgeted Positions records into this file: "
    Msg = Msg & Chr(10) & ImportFile
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do it? "
    a% = MsgBox(Msg, 36, "Confirm")
    
    If a% <> 6 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    If (Dir(ImportFile)) <> "" Then Kill ImportFile

    FileCopy xlsFileTmp, ImportFile
    
    'SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = "SELECT HRJOBBUD.*, JB_STATUS FROM HRJOBBUD LEFT JOIN HRJOB ON HRJOBBUD.JG_CODE = HRJOB.JB_CODE WHERE (1=1) "
    If Len(glbWFCUserSecList) > 0 Then 'Ticket #28254 Franks 04/05/2016
        SQLQ = SQLQ & " AND JG_SECTION IN " & glbWFCUserSecList & " "
    End If
    SQLQ = SQLQ & "ORDER BY JG_YEAR, JG_CODE, JG_CURRENT DESC "
    
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        MsgBox "There is no record in HRJOBBUD table."
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
        
    xRow = 2
    xNum = 0
    xRows = rs.RecordCount
    Do While Not rs.EOF
        MDIMain.panHelp(0).FloodPercent = (xNum / xRows) * 100
        DoEvents
        exSheet.Cells(xRow, 1) = rs("JG_YEAR")
        exSheet.Cells(xRow, 2) = rs("JG_SECTION")
        exSheet.Cells(xRow, 3) = rs("JG_CODE")
        If Not IsNull(rs("JG_CODE")) Then
            exSheet.Cells(xRow, 4) = getPosDesc(rs("JG_CODE"))
        End If
        exSheet.Cells(xRow, 5) = rs("JB_STATUS") 'Ticket #27774 Franks 11/18/2015
        exSheet.Cells(xRow, 5 + 1) = rs("JG_DIV")
        If Not IsNull(rs("JG_DIV")) Then
            exSheet.Cells(xRow, 6 + 1) = getRegionFromDiv(rs("JG_DIV"))
        End If
        exSheet.Cells(xRow, 7 + 1) = rs("JG_DEPTNO")
        exSheet.Cells(xRow, 8 + 1) = rs("JG_GLNO")
        exSheet.Cells(xRow, 9 + 1) = rs("JG_BUDGNBR")
        exSheet.Cells(xRow, 10 + 1) = rs("JG_FTENUM")
        exSheet.Cells(xRow, 11 + 1) = rs("JG_FTEHRS")
        xNum = xNum + 1
        xRow = xRow + 1
        rs.MoveNext
    Loop
    rs.Close
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = ""
    
    Screen.MousePointer = vbDefault
    
    MsgBox "   Finished!   "


End Sub

Private Sub cmdImp_Click()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rs As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim ImportFile
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String
Dim xNum As Integer
Dim xRows As Long
Dim xRow As Long
Dim xEmpnbr
Dim xFlag As Boolean
Dim xYear, xPlant, xPos As String, xDIV, xBUnit, xDept As String, xGL As String, xBudPos, xFte, xFTEHrs, xUptMsg
Dim xTmp

    If Not gSec_Upd_BudgetedPos Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    'Budgeted Position Import.xls or Budgeted Position Import.xlsx
    'check file name
    ImportFile = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Budgeted Position Import.xls"
    If Dir(ImportFile) = "" Then
      MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
      Exit Sub
    End If
    
    Msg = "This program will load Budgeted Positions from the file: "
    Msg = Msg & Chr(10) & ImportFile
    Msg = Msg & Chr(10) & Chr(10) & "Note: click 'Count Budgeted Positions' button to get 'Positions Filled'. "
    Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do it? "
    a% = MsgBox(Msg, 36, "Confirm")
    
    If a% <> 6 Then Exit Sub
    
    'delete wroking table
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
    xRows = getRows(exSheet)
    
    xFlag = True
    xYear = exSheet.Cells(2, 1) 'exSheet.Cells(2, 31)
    If Len(xYear) < 1 Then
        'MsgBox ("Year is a required field")
        xFlag = False
    Else
        If Not IsNumeric(xYear) Then
            'MsgBox "Year must be numeric."
            xFlag = False
        End If
        If Not Len(xYear) = 4 Then
            'MsgBox "Invalid Year."
            xFlag = False
        End If
    End If
    If xFlag = False Then
        MsgBox "Invalid Year in Cell(A2)"
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Exit Sub
    End If
    
    
    For xRow = 2 To xRows
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xUptMsg = ""
        
        xYear = exSheet.Cells(xRow, 1)
        xFlag = True
        If Len(xYear) < 1 Then
            'MsgBox ("Year is a required field")
            xFlag = False
        Else
            If Not IsNumeric(xYear) Then
                'MsgBox "Year must be numeric."
                xFlag = False
            End If
            If Not Len(xYear) = 4 Then
                'MsgBox "Invalid Year."
                xFlag = False
            End If
        End If
        If xFlag = False Then
            xUptMsg = "Invalid Year": GoTo Error_Upt
        End If
        
        Call getFromToDate(xYear)

        xPlant = exSheet.Cells(xRow, 2)
        If Len(xPlant) < 1 Then
            xUptMsg = "No Plant Code": GoTo Error_Upt
        End If
        xTmp = GetTABLDesc("EDSE", xPlant)
        If Len(xTmp) = 0 Then
            xUptMsg = "Invalid Plant Code": GoTo Error_Upt
        End If
        xPos = exSheet.Cells(xRow, 3)
        If Len(xPos) < 1 Then
            xUptMsg = "No Position Code": GoTo Error_Upt
        End If
        xTmp = getPosDesc(xPos)
        If Len(xTmp) = 0 Then
            xUptMsg = "Invalid Position Code": GoTo Error_Upt
        End If
        xDIV = exSheet.Cells(xRow, 5 + 1)
        If Len(xDIV) < 1 Then
            xUptMsg = "No Division Code": GoTo Error_Upt
        End If
        xTmp = getDivDescPub(xDIV)
        If Len(xTmp) = 0 Then
            xUptMsg = "Invalid Division Code": GoTo Error_Upt
        End If
        
        'Position/Division must be setup in HRJOB. If not, display a message saying "This Position has not been assigned to this Division". Dead stop.
        SQLQ = "SELECT JB_CODE FROM HRJOB WHERE JB_CODE='" & Trim(xPos) & "' "
        SQLQ = SQLQ & "AND JB_DIV = '" & xDIV & "' "
        If rsJOB.State <> 0 Then rsJOB.Close
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If rsJOB.EOF Then
            xUptMsg = "This Position has not been assigned to this Division.": GoTo Error_Upt
        End If
        rsJOB.Close

        xBUnit = exSheet.Cells(xRow, 6 + 1)
        If Len(xBUnit) > 0 Then
            xTmp = GetTABLDesc("EDRG", xBUnit)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid Business Unit Code": GoTo Error_Upt
            End If
        End If
        xDept = exSheet.Cells(xRow, 7 + 1)
        If Len(xDept) > 0 Then
            xTmp = getDeptDescPub(xDept)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid Department Code": GoTo Error_Upt
            End If
        End If
        xGL = exSheet.Cells(xRow, 8 + 1)
        If Len(xGL) > 0 Then
            xTmp = getGLDesc(xGL)
            If Len(xTmp) = 0 Then
                xUptMsg = "Invalid GL Code": GoTo Error_Upt
            End If
        End If
        xBudPos = exSheet.Cells(xRow, 9 + 1)
        If Len(xBudPos) < 0 Then xBudPos = ""
        If Len(xBudPos) > 0 Then
            If Not IsNumeric(xBudPos) Then
                xUptMsg = "Invalid Budgeted Position": GoTo Error_Upt
            End If
        End If
        xFte = exSheet.Cells(xRow, 10 + 1)
        If Len(xFte) < 0 Then xFte = ""
        If Len(xFte) > 0 Then
            If Not IsNumeric(xFte) Then
                xUptMsg = "Invalid FTE #": GoTo Error_Upt
            End If
        End If
        xFTEHrs = exSheet.Cells(xRow, 11 + 1)
        If Len(xFTEHrs) < 0 Then xFTEHrs = ""
        If Len(xFTEHrs) > 0 Then
            If Not IsNumeric(xFTEHrs) Then
                xUptMsg = "Invalid FTE Hours/Year": GoTo Error_Upt
            End If
        End If
Error_Upt:
        If Len(xUptMsg) > 0 Then
            exSheet.Cells(xRow, 12 + 1) = xUptMsg
            GoTo Next_Rec
        End If
        
        '''check duplicate record
        ''If modISDupBudget(xPos, xPlant, xDiv, xDept, xGL, 0, fglbFDate, fglbTDate, fglbEDate) Then
        ''    exSheet.Cells(xRow, 12) = "Duplicate record."
        ''    GoTo Next_Rec
        ''End If
        '''xYear, xPlant, xPos, xDiv, xBUnit, xDept, xGL, xBudPos, xFTE, FTEHrs, xUptMsg
        '''update this record
        
        SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xPos & "' "
        SQLQ = SQLQ & "AND JG_YEAR = " & xYear & " "
        SQLQ = SQLQ & "AND JG_SECTION = '" & xPlant & "' "
        If Len(xDIV) > 0 Then
            SQLQ = SQLQ & "AND JG_DIV = '" & xDIV & "' "
        End If
        If Len(xDept) > 0 Then
            SQLQ = SQLQ & "AND JG_DEPTNO = '" & xDept & "' "
        End If
        If Len(xGL) > 0 Then
            SQLQ = SQLQ & "AND JG_GLNO = '" & xGL & "' "
        End If
        If IsDate(fglbFDate) Then
            SQLQ = SQLQ & "AND JG_FRDATE = " & Date_SQL(fglbFDate) & " "
        End If
        If IsDate(fglbTDate) Then
            SQLQ = SQLQ & "AND JG_TODATE = " & Date_SQL(fglbTDate) & " "
        End If
        If IsDate(fglbEDate) Then
            SQLQ = SQLQ & "AND JG_EFDATE = " & Date_SQL(fglbEDate) & " "
        End If

        
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            rs.AddNew
            rs("JG_COMPNO") = "001"
            rs("JG_CODE") = xPos
        End If
            'rs("JG_BUDPOSNBR") = Null
            If Len(xDIV) > 0 Then rs("JG_DIV") = xDIV
            If Len(xDept) = 0 Then rs("JG_DEPTNO") = xDept
            If Len(xGL) = 0 Then rs("JG_GLNO") = xGL
            'rs("JG_NBRFIL") = rsFrom("JG_NBRFIL")
            If Len(xBudPos) > 0 And IsNumeric(xBudPos) Then
                rs("JG_BUDGNBR") = xBudPos
            End If
            If Len(xFte) > 0 Then
                rs("JG_FTENUM") = xFte
            End If
            'rs("JG_FTENUMFILL") = rsFrom("JG_FTENUMFILL")
            'rs("JG_FTENUMVACN") = rsFrom("JG_FTENUMVACN")
            If Len(xFTEHrs) > 0 Then
                rs("JG_FTEHRS") = xFTEHrs
            End If
            'rs("JG_FTETOTHR") = rsFrom("JG_FTETOTHR")
            rs("JG_LDATE") = Date
            rs("JG_LTIME") = Time$
            rs("JG_LUSER") = glbUserID
            rs("JG_SECTION") = xPlant 'rsFrom("JG_SECTION")
            rs("JG_YEAR") = xYear 'rsFrom("JG_YEAR")
            rs("JG_FRDATE") = fglbFDate ' rsFrom("JG_FRDATE")
            rs("JG_TODATE") = fglbTDate 'rsFrom("JG_TODATE")
            rs("JG_EFDATE") = fglbEDate 'rsFrom("JG_EFDATE")
            'rs("JG_JREASON") = "" 'rsFrom("JG_JREASON")
            'rs("JG_VACANCY_POS") = rsFrom("JG_VACANCY_POS")
            'rs("JG_CURRENT") = rsFrom("JG_CURRENT")
            rs.Update
        'End If
        rs.Close

        'add to the wrk table
        SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND TT_EMPNBR = 9999 "
        SQLQ = SQLQ & "AND TT_JOB = '" & xPos & "' "
        SQLQ = SQLQ & "AND TT_CHAR10 = '" & xPlant & "' "
        
        If rsWRK.State <> 0 Then rsWRK.Close
        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWRK.EOF Then
            rsWRK.AddNew
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK("TT_EMPNBR") = 9999
            rsWRK("TT_JOB") = xPos
            rsWRK("TT_CHAR10") = xPlant
            rsWRK.Update
        End If
        rsWRK.Close

Next_Rec:

    Next
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    'setup current flag
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xRow = 0
    If Not rsWRK.EOF Then
        xRows = rsWRK.RecordCount
    End If
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        xRow = xRow + 1
        DoEvents
        Call setCurrentFlag(rsWRK("TT_JOB"), rsWRK("TT_CHAR10"))
        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    
    Screen.MousePointer = vbDefault
    
    'MsgBox "   Finished!   "
    Unload Me
End Sub

Private Function getRows(exSheet As Object)
Dim X
X = 1
Do While True
    If exSheet.Cells(X, 1) = "" Then
        Exit Do
    Else
        X = X + 1
    End If
Loop
getRows = X - 1
End Function

Public Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_BudgetedPos Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo Edit_Err


fglbEditMode% = True

RDept = clpDept

clpDiv.SetFocus

Exit Sub

Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRJOBSKL", "Edit")
Call RollBack   '15June99 js
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String
Dim xYear

If Not gSec_Upd_BudgetedPos Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
fglbNew = True
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me, rsDATA)
rsDATA.AddNew

'Data1.Recordset.AddNew
fglbEditMode% = True
lblCNum.Caption = "001"
lblPositions.Caption = glbPos$
lblSections.Caption = glbJobSection

If Mid(lblPosition.Caption, 5, 3) = "IND" Then  'Ticket #30358 Franks 07/13/2017 - leave these as blank for Independent Contractor Positions
    medFTEHrs.Text = 0
Else
    medFTEHrs.Text = 2080 'Ticket #29005 Franks 08/02/2016
End If
'Ticket #29183 Franks 09/12/2016 - begin
If month(Date) = 11 Or month(Date) = 12 Then
    xYear = Year(Date) + 1
Else
    xYear = Year(Date)
End If
txtYear.Text = xYear
Call getDateRange(xYear)
clpDiv.Text = Left(glbPos$, 4)
txtNoPos.Text = 1
medFTENum.Text = 1
chkCurrent.Value = True
'Ticket #29183 Franks 09/12/2016 - end
'clpDept.Enabled = True 'Ticket #29005 Franks 08/02/2016

'clpDiv.SetFocus
txtYear.SetFocus

RDept = ""
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBBUD", "Add")
Call RollBack
End Sub

Public Sub cmdOK_Click()
On Error GoTo OK_Err

If Not chkBudgetPos() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Call setCurrentFlag(glbPos, glbJobSection)
Call mod_Upd_Pos_Budget_WFC(glbPos, glbJobSection)
Data1.Refresh
Call Display_Value

fglbNew = False
'Call ST_UPD_MODE(False)
Call SET_UP_MODE
fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBBUD", "Update")
Call RollBack   '15June99 js
Unload Me

End Sub

Private Sub CalcCurFlagForAll()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPos
Dim xJobSection
Dim xNum, xRows

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).Caption = ""
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
    
SQLQ = "SELECT * FROM HRJOB ORDER BY JB_CODE "
If rs.State <> 0 Then rs.Close
rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
xNum = 0
If Not rs.EOF Then
    xRows = rs.RecordCount
End If
Do While Not rs.EOF
    MDIMain.panHelp(0).FloodPercent = (xNum / xRows) * 100
    xNum = xNum + 1
    DoEvents
        
    xJobSection = rs("JB_SECTION")
    Call setCurrentFlag(rs("JB_CODE"), xJobSection)
    rs.MoveNext
Loop
rs.Close

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).Caption = ""
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Screen.MousePointer = DEFAULT

End Sub

Private Sub CalcCurFlagForPlant(xJobSection)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPos

If Len(xJobSection) > 0 Then
    SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & xJobSection & "' ORDER BY JB_CODE "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rs.EOF
        Call setCurrentFlag(rs("JB_CODE"), xJobSection)
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

Private Sub setCurrentFlag(xPos, xJobSection)
Dim rsMain As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "UPDATE HRJOBBUD SET JG_CURRENT = 0 WHERE JG_CODE = '" & xPos & "' "
    SQLQ = SQLQ & "AND JG_SECTION = '" & xJobSection & "' "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "SELECT JG_CODE,JG_SECTION,JG_DIV,JG_DEPTNO FROM HRJOBBUD WHERE JG_CODE = '" & xPos & "' "
    SQLQ = SQLQ & "AND JG_SECTION = '" & xJobSection & "' "
    SQLQ = SQLQ & "ORDER BY JG_CODE,JG_SECTION,JG_DIV,JG_DEPTNO  "
    If rsMain.State <> 0 Then rsMain.Close
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsMain.EOF
        SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xPos & "' "
        SQLQ = SQLQ & "AND JG_SECTION = '" & xJobSection & "' "
        If Not IsNull(rsMain("JG_DIV")) Then
            If Len(rsMain("JG_DIV")) > 0 Then
                SQLQ = SQLQ & "AND JG_DIV = '" & rsMain("JG_DIV") & "' "
            End If
        End If
        If Not IsNull(rsMain("JG_DEPTNO")) Then
            If Len(rsMain("JG_DEPTNO")) > 0 Then
                SQLQ = SQLQ & "AND JG_DEPTNO = '" & rsMain("JG_DEPTNO") & "' "
            End If
        End If
        SQLQ = SQLQ & "ORDER BY JG_YEAR DESC,JG_EFDATE DESC "
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            rs("JG_CURRENT") = 1
            rs.Update
        End If
        rs.Close
        
        rsMain.MoveNext
    Loop
    rsMain.Close

''    SQLQ = "UPDATE HRJOBBUD SET JG_CURRENT = 0 WHERE JG_CODE = '" & xPos & "' "
''    SQLQ = SQLQ & "AND JG_SECTION = '" & xJobSection & "' "
''    If Not IsMissing(xJobDept) Then
''        If Len(xJobDept) > 0 Then
''            SQLQ = SQLQ & "AND JG_DEPTNO = '" & xJobDept & "' "
''        End If
''    End If
''    gdbAdoIhr001.Execute SQLQ
''    SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xPos & "' "
''    SQLQ = SQLQ & "AND JG_SECTION = '" & xJobSection & "' "
''    If Not IsMissing(xJobDept) Then
''        If Len(xJobDept) > 0 Then
''            SQLQ = SQLQ & "AND JG_DEPTNO = '" & xJobDept & "' "
''        End If
''    End If
''    SQLQ = SQLQ & "ORDER BY JG_YEAR DESC,JG_EFDATE DESC "
''    If rs.State <> 0 Then rs.Close
''    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''    If Not rs.EOF Then
''        rs("JG_CURRENT") = 1
''        rs.Update
''    End If
''    rs.Close
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

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

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMPOSBUDGETWFC"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%

On Error GoTo FLErr

glbOnTop = "FRMPOSBUDGETWFC" ' "frmPosBudgetWFC"

Screen.MousePointer = HOURGLASS
'If glbPos = "" Then frmJOBS.Show 1 'frmJOBSWFC.Show 1 '
If glbPos = "" Then frmJOBSWFC.Show 1 '
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
lblSec.Caption = glbJobSection
glbJobSectionDesc = getCodeDesc("EDSE", glbJobSection)
lblSecDesc.Caption = glbJobSectionDesc

Me.Caption = "Budgeted Positions - " & lblPosition

Data1.ConnectionString = glbAdoIHRDB
'Call CR_Lgr_Snap

If Not EERetrieve() Then
    Exit Sub        '  modGet it sets fglbRecords
End If
lblDiv.Caption = lStr(lblDiv)
lblDept.Caption = lStr(lblDept)
lblGLNo.Caption = lStr(lblGLNo)
lblTitle(1).Caption = lStr("Section")
lblTitle(18).Caption = lStr("Section")
lblTitle(19).Caption = lStr("Section")

Call INI_Controls(Me)
Call Display_Value

Call SET_UP_MODE

frmCopyOne.Top = frmCopy.Top
frmCopyOne.Left = 4080
frmCopyPlant.Top = 4320 'frmCopy.Top
frmCopyPlant.Left = 4080
frmCopyPlant.Height = 4335
frmDelete.Top = frmCopy.Top
frmDelete.Left = 4080

chkCopy.Visible = False 'Ticket #26507 Franks 01/14/2015

'Ticket #25911 Franks 01/28/2015
'3. optional Department and GL Code casuse the program very complicate. Talk this with Jerry, he let Frank to disable these two fields for now, if WFC really need these then we will make these enable
lblDept.Enabled = False
clpDept.Enabled = False
lblGLNo.Enabled = False
clpGLNum.Enabled = False

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Budgeted Positions", "Select")
Call RollBack   '15June99 js


End Sub

Private Function getJobSec(xCode)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HRJOB WHERE "
    
    SQLQ = "SELECT JB_CODE,JB_SECTION FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        If Not IsNull(rs("JB_SECTION")) Then
            retVal = rs("JB_SECTION")
        End If
    End If
    rs.Close
    
    getJobSec = retVal
End Function

Public Function EERetrieve() 'StrPos$)
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr

glbJobSection = getJobSec(glbPos)

' out or left join query not updateable - so do straight.
SQLQ$ = "SELECT * FROM HRJOBBUD "
SQLQ$ = SQLQ$ & "WHERE JG_CODE = '" & glbPos & "' "
If Len(glbJobSection) = 0 Then
    SQLQ$ = SQLQ$ & "AND JG_SECTION IS NULL "
Else
    SQLQ$ = SQLQ$ & "AND JG_SECTION = '" & glbJobSection & "' "
End If
SQLQ$ = SQLQ$ & "ORDER BY JG_CODE, JG_EFDATE DESC"

Data1.RecordSource = SQLQ$
Data1.Refresh

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
lblSec.Caption = glbJobSection
glbJobSectionDesc = getCodeDesc("EDSE", glbJobSection)
lblSecDesc.Caption = glbJobSectionDesc

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    fglbRecords% = False
    cmdModify.Enabled = False       'Laura jan 06, 1998
Else
    fglbRecords% = True
End If
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

Private Sub imgPosFilled_Click()
Dim xMsg As String
    'xMsg = getBugPosEmpList(glbPos, glbJobSection, clpDiv.Text, clpDept.Text, clpGLNum.Text)
    'MsgBox xMsg, , "Positions Filled Employee List"
    Call getBugPosEmpList(glbPos, glbJobSection, clpDiv.Text, clpDept.Text, clpGLNum.Text)
End Sub

Private Sub lblPositions_Change()
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
End Sub

Private Sub lblSections_Change()
lblSec.Caption = glbJobSection
glbJobSectionDesc = getCodeDesc("EDSE", glbJobSection)
lblSecDesc.Caption = glbJobSectionDesc
End Sub

Private Sub medFTEHrs_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medFTENum_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Dept_GL()
Dim snapDepts As New ADODB.Recordset
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ
If Not cmdOK.Enabled Then RDept = clpDept
If Len(clpDept.Text) > 0 Then

    SQLQ = "Select DISTINCT HRDEPT.DF_NBR, HRDEPT.DF_NAME, HRDEPT.DF_GLNO from HRDEPT"
    SQLQ = SQLQ & " Where " & glbSeleDept
    SQLQ = SQLQ & " AND DF_NBR = '" & clpDept.Text & "'"
    If glbOracle Then
        SQLQ = SQLQ & " ORDER BY DF_NAME "
    Else
        SQLQ = SQLQ & " ORDER BY [DF_NAME] "
    End If
    If snapDepts.State <> 0 Then snapDepts.Close
    snapDepts.Open SQLQ, gdbAdoIhr001, adOpenStatic

    If Not snapDepts.EOF Then
        RGLNum = snapDepts("DF_GLNO")
        If RDept <> clpDept Then
            If IsNull(RGLNum) Then
                RGLNum = ""
                'txtGLNum = ""
            Else
                Msg$ = "Do you want the associated G/L #?"
                Title$ = "info:HR"
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDYES Then clpGLNum = RGLNum
            End If
            RDept = clpDept

        End If
    End If
End If
End Sub

Private Sub CR_Lgr_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Job_Err

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT * FROM HRGL "

If LGR_snap.State <> 0 Then LGR_snap.Close
LGR_snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Descriptions", "HRGL", "SELECT")
Call RollBack '21June99 js

End Sub

Private Function chkBudgetPos()
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String, Msg As String, dd#, PID&, xDIV, xDeptno$, xGLNO$, xPosCtrl

chkBudgetPos = False

On Error GoTo chkBudgetPos_Err

If Len(txtYear.Text) < 1 Then
    MsgBox ("Year is a required field")
    txtYear.SetFocus
    Exit Function
Else
    If Not IsNumeric(txtYear.Text) Then
        MsgBox "Year must be numeric."
        txtYear.SetFocus
        Exit Function
    End If
    If Not Len(txtYear.Text) = 4 Then
        MsgBox "Invalid Year."
        txtYear.SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(0).Text) < 1 Then
    MsgBox "From Date is a required field"
    dlpDate(0).SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Invalid Date"
        dlpDate(0).SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(1).Text) < 1 Then
    MsgBox "To Date is a required field"
    dlpDate(1).SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Invalid Date"
        dlpDate(1).SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(2).Text) < 1 Then
    MsgBox "Effective Date is a required field"
    dlpDate(2).SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate(2).Text) Then
        MsgBox "Invalid Date"
        dlpDate(2).SetFocus
        Exit Function
    End If
    If CVDate((dlpDate(2).Text)) < CVDate((dlpDate(0).Text)) Then
        MsgBox "Effective Date must be within the Date Range"
        dlpDate(2).SetFocus
        Exit Function
    End If
    If CVDate((dlpDate(2).Text)) > CVDate((dlpDate(1).Text)) Then
        MsgBox "Effective Date must be within the Date Range"
        dlpDate(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(0).Text) < 1 Then
    If IsDate(dlpDate(0).Text) And IsDate(dlpDate(2).Text) Then
        If Not (CVDate((dlpDate(0).Text)) = CVDate((dlpDate(2).Text))) Then
            MsgBox ("Reason is a required field")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
Else
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox ("Reason must be valid")
        clpCode(0).SetFocus
        Exit Function
    End If
End If

If Len(clpDiv.Text) < 1 Then
    MsgBox lStr("Division is a required field")
    clpDiv.SetFocus
    Exit Function
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("Division must be valid")
        clpDiv.SetFocus
        Exit Function
    End If
End If

If Len(clpDept) < 1 Then
'    MsgBox lStr("Department is a required field")
'    clpDept.SetFocus
'    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        If clpDept.Enabled Then
        MsgBox lStr("Department must be valid")
        clpDept.SetFocus
        Exit Function
        End If
    End If
End If

If clpGLNum.Enabled Then
    If clpGLNum.Caption = "Unassigned" Then
        MsgBox lStr("G/L Code must be valid")
        clpGLNum.SetFocus
        Exit Function
    End If
End If

If IsNull(rsDATA("JG_ID")) Then
    PID& = 0
Else
    PID& = rsDATA("JG_ID")
End If
xDIV = clpDiv
xDeptno$ = clpDept
xGLNO$ = clpGLNum

'Position/Division must be setup in HRJOB. If not, display a message saying "This Position has not been assigned to this Division". Dead stop.
SQLQ = "SELECT JB_CODE FROM HRJOB WHERE JB_CODE='" & Trim(glbPos$) & "' "
SQLQ = SQLQ & "AND JB_DIV = '" & clpDiv.Text & "' "
rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsJOB.EOF Then
    MsgBox "This Position has not been assigned to this Division."
    clpDiv.SetFocus
    Exit Function
End If
rsJOB.Close

'If Not glbWHSCC Then
'    If modISDupBudgetPosCtrl(glbPos$, xPosCtrl, PID&) Then
'        MsgBox "Position Control # must be unique"
'        clpDiv.SetFocus
'        Exit Function
'    End If
'Else
    If modISDupBudget(glbPos, glbJobSection, xDIV, xDeptno$, xGLNO$, PID&, dlpDate(0).Text, dlpDate(1).Text, dlpDate(2).Text) Then
        'If Len(xGLNO$) > 0 Then
        '    If Len(xDiv) > 0 Then
        '        MsgBox lStr("[Division]") & " + " & lStr("[Department]") & " + " & lStr("[G/L Code]") & " must be unique"
        '    Else
        '        MsgBox lStr("[Department]") & " + " & lStr("[G/L Code]") & " must be unique"
        '    End If
        'Else
        '    If Len(xDiv) > 0 Then
        '        MsgBox lStr("[Division]") & " + " & lStr("[Department]") & " must be unique"
        '    Else
        '        MsgBox lStr("[Department]") & " must be unique"
        '    End If
        'End If
        'clpDiv.SetFocus
        Msg = "Duplicate record found."
        Msg = Msg & Chr(10) & lStr("[Position] + ") & lStr("[Section]") & " + " & lStr("[Division]") & " + [Department] + [GL]" & " + [From Date] + [To Date] + [Effective Date] must be unique"
        'MsgBox lStr("[Position] + ") & lStr("[Section]") & " + " & lStr("[Division]") & " + [From Date] + [To Date] + [Effective Date] must be unique"
        MsgBox Msg
        txtYear.SetFocus
        Exit Function
    End If
'End If
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
Private Function modISDupBudget(Pos$, xSec, xDIV, DeptNo$, GLNO$, ID&, xFDate, xTDate, xEDate)
Dim SQLQ
Dim snapBudget As New ADODB.Recordset

modISDupBudget = True

On Error GoTo modISDupBudget_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & Pos$ & "' "
SQLQ = SQLQ & "AND JG_SECTION = '" & xSec & "' "
SQLQ = SQLQ & "AND JG_DIV = '" & xDIV & "' "
If IsDate(xFDate) Then
    SQLQ = SQLQ & "AND JG_FRDATE = " & Date_SQL(xFDate) & " "
End If
If IsDate(xTDate) Then
    SQLQ = SQLQ & "AND JG_TODATE = " & Date_SQL(xTDate) & " "
End If
If IsDate(xEDate) Then
    SQLQ = SQLQ & "AND JG_EFDATE = " & Date_SQL(xEDate) & " "
End If
If Len(DeptNo$) > 0 Then
    SQLQ = SQLQ & "AND JG_DEPTNO = '" & DeptNo$ & "' "
End If
If Len(GLNO$) > 0 Then
    SQLQ = SQLQ & "AND JG_GLNO = '" & GLNO$ & "' "
End If
If fglbNew Or ID& = 0 Then
Else
SQLQ = SQLQ & "AND JG_ID <> " & ID& & " "
End If
If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapBudget.BOF And snapBudget.EOF Then
    modISDupBudget = False
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

Private Sub OptCopyRec_Click(Index As Integer)
If OptCopyRec(0).Value Then
    frmCopyOne.Visible = True
    frmCopyPlant.Visible = False
Else
    frmCopyOne.Visible = False
    frmCopyPlant.Visible = True
    Call ScreensetupForCopyPlant(Index)
End If
End Sub
Private Sub ScreensetupForCopyPlant(xInx)
    If xInx = 1 Then
        lblTitle(11).Enabled = False
        lblTitle(12).Enabled = False
        lblTitle(16).Enabled = False
        lblTitle(17).Enabled = False
        clpDivF.Enabled = False
        clpDivT.Enabled = False
        clpDeptF.Enabled = False
        clpDeptT.Enabled = False
        clpDivF.Text = ""
        clpDivT.Text = ""
        clpDeptF.Text = ""
        clpDeptT.Text = ""
    End If
    If xInx = 2 Then
        lblTitle(11).Enabled = True
        lblTitle(12).Enabled = False
        lblTitle(16).Enabled = True
        lblTitle(17).Enabled = False
        clpDivF.Enabled = True
        clpDivT.Enabled = True
        clpDeptF.Enabled = False
        clpDeptT.Enabled = False
        clpDivF.Text = ""
        clpDivT.Text = ""
        clpDeptF.Text = ""
        clpDeptT.Text = ""
    End If
    If xInx = 3 Then
        lblTitle(11).Enabled = True
        lblTitle(12).Enabled = True
        lblTitle(16).Enabled = True
        lblTitle(17).Enabled = True
        clpDivF.Enabled = True
        clpDivT.Enabled = True
        clpDeptF.Enabled = True
        clpDeptT.Enabled = True
        clpDivF.Text = ""
        clpDivT.Text = ""
        clpDeptF.Text = ""
        clpDeptT.Text = ""
    End If
    
End Sub

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
    SQLQ = "SELECT * FROM HRJOBBUD "
    SQLQ = SQLQ & "WHERE JG_ID = " & Data1.Recordset!JG_ID
    SQLQ = SQLQ & " ORDER BY JG_CODE,JG_EFDATE DESC"

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    lblID = rsDATA!JG_ID
    Call Set_Control("R", Me, rsDATA)
End If
    Call SET_UP_MODE
End Sub

Private Sub txtNoPos_LostFocus()
If fglbNew Then 'Ticket #29005 Franks 08/02/2016
    If Len(medFTENum.Text) = 0 Then
        medFTENum.Text = txtNoPos.Text
    End If
End If
End Sub

Private Sub txtYear_LostFocus()
    Call getDateRange(txtYear.Text)
End Sub
Private Sub getDateRange(xYear)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
If fglbNew Then
    If IsNumeric(xYear) Then
        If Len(xYear) = 4 Then
            SQLQ = "SELECT * FROM HRPARCO"
            If rs.State <> 0 Then rs.Close
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rs.EOF Then
                If Not IsNull(rs("PC_TDATE")) Then
                    I = xYear - Year(rs("PC_TDATE"))
                    dlpDate(0).Text = DateAdd("YYYY", I, rs("PC_FDATE"))
                    dlpDate(1).Text = DateAdd("YYYY", I, rs("PC_TDATE"))
                    dlpDate(2).Text = dlpDate(0).Text
                End If
            End If
            rs.Close
        End If
    End If
End If
End Sub

Private Sub getFromToDate(xYear)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
    fglbFDate = ""
    fglbTDate = ""
    fglbEDate = ""
    If IsNumeric(xYear) Then
        If Len(xYear) = 4 Then
            SQLQ = "SELECT * FROM HRPARCO"
            If rs.State <> 0 Then rs.Close
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rs.EOF Then
                If Not IsNull(rs("PC_TDATE")) Then
                    I = xYear - Year(rs("PC_TDATE"))
                    fglbFDate = DateAdd("YYYY", I, rs("PC_FDATE"))
                    fglbTDate = DateAdd("YYYY", I, rs("PC_TDATE"))
                    fglbEDate = fglbFDate
                End If
            End If
            rs.Close
        End If
    End If

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        ' out or left join query not updateable - so do straight.
        SQLQ$ = "SELECT * FROM HRJOBBUD "
        SQLQ$ = SQLQ$ & "WHERE JG_CODE = '" & glbPos$ & "' "
        If Len(glbJobSection) = 0 Then
            SQLQ$ = SQLQ$ & "AND JG_SECTION IS NULL "
        Else
            SQLQ$ = SQLQ$ & "AND JG_SECTION = '" & glbJobSection & "' "
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X%
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value

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
UpdateRight = gSec_Upd_BudgetedPos
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
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

txtYear.Enabled = TF
chkCurrent.Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
dlpDate(2).Enabled = TF
clpCode(0).Enabled = TF
clpDiv.Enabled = TF
clpDept.Enabled = False ' TF
clpGLNum.Enabled = False 'TF
txtNoPos.Enabled = TF
medFTENum.Enabled = TF
medFTEHrs.Enabled = TF
'medVacaPos.Enabled = TF

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

Private Function getRegionFromDiv(xDIV)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDIV & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        If Not IsNull(rs("DV_REGION")) Then
            retVal = rs("DV_REGION")
        End If
    End If
    rs.Close
    getRegionFromDiv = retVal
End Function

Private Function getBugPosEmpList(glbPos, glbJobSection, xlocDiv, xlocDept, xlocGLNum)
Dim snapJobCount As New ADODB.Recordset
Dim rsHRJOB As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDIV, xDeptno, xGLNO, xPosCtrl
Dim xSec, xCunt
Dim xBudgNo, xVacantNo, I
Dim retVal
    retVal = ""
    
    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
    SQLQ = SQLQ & "AND JG_CURRENT = 1 "
    If Not Len(glbPos) = 0 Then SQLQ = SQLQ & "AND JG_CODE = '" & glbPos & "' "
    If Not Len(glbJobSection) = 0 Then SQLQ = SQLQ & "AND JG_SECTION = '" & glbJobSection & "' "
    If Not Len(xlocDiv) = 0 Then SQLQ = SQLQ & "AND JG_DIV = '" & xlocDiv & "' "
    If Not Len(xlocDept) = 0 Then SQLQ = SQLQ & "AND JG_DEPTNO = '" & xlocDept & "' "
    If Not Len(xlocGLNum) = 0 Then SQLQ = SQLQ & "AND clpGLNum = '" & xlocGLNum & "' "
    
    xCunt = 0
    xBudgNo = 0
    
    If snapBudget.State <> 0 Then snapBudget.Close
    snapBudget.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not snapBudget.EOF Then
        
        'Ticket #27820 Franks 11/26/2015
        If Not IsNull(snapBudget("JG_BUDGNBR")) Then
            xBudgNo = snapBudget("JG_BUDGNBR")
        End If
        
        If Not IsNull(snapBudget("JG_SECTION")) Then
            xSec = snapBudget("JG_SECTION")
        End If
        If Not IsNull(snapBudget("JG_DIV")) Then
            xDIV = snapBudget("JG_DIV")
        End If
        If Not IsNull(snapBudget("JG_DEPTNO")) Then
            xDeptno = snapBudget("JG_DEPTNO")
        End If
        If Not IsNull(snapBudget("JG_GLNO")) Then
            xGLNO = snapBudget("JG_GLNO")
        End If
        'Position filled
    
        SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,ED_EMPNBR,ED_SURNAME,ED_FNAME "
        If Len(xSec) > 0 Then SQLQ = SQLQ & ",ED_SECTION "
        If Len(xDIV) > 0 Then SQLQ = SQLQ & ",ED_DIV "
        If Len(xDeptno) > 0 Then SQLQ = SQLQ & ",ED_DEPTNO "
        If Len(xGLNO) > 0 Then SQLQ = SQLQ & ",ED_GLNO "
        'SQLQ = SQLQ & "COUNT(HR_JOB_HISTORY.JH_EMPNBR) AS NoPosFilled  "
        SQLQ = SQLQ & "FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & "INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & glbPos & "' "
        SQLQ = SQLQ & "AND NOT HREMP.ED_EMP = 'RET' " 'exclude employees the RET Employment Status
        SQLQ = SQLQ & "AND NOT (ED_FNAME LIKE '%(Deceased)%') " 'not include Death employees
        If Len(xSec) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_SECTION = '" & xSec & "' "
        If Len(xDIV) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDIV & "' "
        If Len(xDeptno) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
        If Len(xGLNO) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
        SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME"
        'If Len(xSec) > 0 Then SQLQ = SQLQ & ",HREMP.ED_SECTION "
        'If Len(xDiv) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DIV "
        'If Len(xDeptno) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DEPTNO "
        'If Len(xGLNO) > 0 Then SQLQ = SQLQ & ",HREMP.ED_GLNO "
    
        snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not snapJobCount.EOF
            'If snapJobCount.RecordCount <= 30 Then
            '    retval = retval & snapJobCount("ED_SURNAME") & "," & snapJobCount("ED_FNAME") & Chr(10)
            'Else
            '    retval = retval & snapJobCount("ED_SURNAME") & "," & snapJobCount("ED_FNAME") & "|"
            'End If
            xCunt = xCunt + 1
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "001"
            rsEListWRK("TT_EMPNBR") = snapJobCount("ED_EMPNBR")
            rsEListWRK("TT_SURNAME") = snapJobCount("ED_SURNAME")
            rsEListWRK("TT_FNAME") = snapJobCount("ED_FNAME")
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
            snapJobCount.MoveNext
        Loop
        snapJobCount.Close
    
    End If
    'rsEListWRK.Close
    
    'Ticket #27820 Franks 11/26/2015 - begin
    xVacantNo = xBudgNo - xCunt
    If xVacantNo > 0 Then
        For I = 1 To xVacantNo
            xCunt = xCunt + 1
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "002"
            rsEListWRK("TT_EMPNBR") = 0 'snapJobCount("ED_EMPNBR")
            rsEListWRK("TT_SURNAME") = "Vacant " '& Trim(Str(I))  'snapJobCount("ED_SURNAME")
            rsEListWRK("TT_FNAME") = "" 'snapJobCount("ED_FNAME")
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
        Next
    End If
    rsEListWRK.Close
    'Ticket #27820 Franks 11/26/2015 -end
    
    getBugPosEmpList = retVal
    
    If xCunt = 0 Then
        MsgBox "No record found."
    Else
        Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "RZEmpList3.rpt"
        Me.vbxCrystal2.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
        Me.vbxCrystal2.Formulas(0) = "rTitle='Employee list for Position " & glbPos & "'"
        Me.vbxCrystal2.Connect = RptODBC_SQL
        'window title if appropriate
        Me.vbxCrystal2.WindowTitle = "Position Employees List Report"
        Me.vbxCrystal2.Destination = 0
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal2.Action = 1
        vbxCrystal2.Reset

    End If
    
End Function
