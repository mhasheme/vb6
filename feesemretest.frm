VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmESEMRETEST 
   Caption         =   "Continuing Education - Retest"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDocImport 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   4920
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   3000
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   2640
         Picture         =   "feesemretest.frx":0000
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuing Education Retest"
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
         Left            =   -120
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   2640
         Picture         =   "feesemretest.frx":014A
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6840
      MaxLength       =   25
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ES_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtCourseHRS 
      Appearance      =   0  'Flat
      DataField       =   "ES_HOURS"
      Height          =   285
      Left            =   7080
      MaxLength       =   5
      TabIndex        =   9
      Tag             =   "11-Number of Scheduled Course Hours "
      Top             =   2400
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feesemretest.frx":0294
      Height          =   1425
      Left            =   120
      OleObjectBlob   =   "feesemretest.frx":02A8
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin INFOHR_Controls.CodeLookup clpEmpCur 
      DataField       =   "ES_EMPCUR"
      Height          =   285
      Left            =   8220
      TabIndex        =   16
      Top             =   2760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.DateLookup dlpRenewal 
      DataField       =   "ES_RENEW"
      Height          =   285
      Left            =   1620
      TabIndex        =   8
      Tag             =   "40-Date when course is to be renewed"
      Top             =   4920
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_RESULTS"
      Height          =   285
      Index           =   3
      Left            =   1620
      TabIndex        =   6
      Tag             =   "00-Results of the Course - Code"
      Top             =   4200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESRT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CONDUCT"
      Height          =   285
      Index           =   2
      Left            =   1620
      TabIndex        =   7
      Tag             =   "00-Organization/Individual Instructing - Code"
      Top             =   4560
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CRSCODE"
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Tag             =   "00-Course Code"
      Top             =   2760
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CTYPE"
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Tag             =   "01-Course Type - Code"
      Top             =   2400
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCT"
      MaxLength       =   8
      Enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox medEECont 
      DataField       =   "ES_TBEMP"
      Height          =   285
      Index           =   0
      Left            =   7080
      TabIndex        =   10
      Tag             =   "20-Amount Employee Paid"
      Top             =   2760
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
      Left            =   7080
      TabIndex        =   11
      Tag             =   "20-Other Expenses Paid"
      Top             =   3120
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
      Left            =   7080
      TabIndex        =   12
      Tag             =   "20-Amount Employer Paid"
      Top             =   3480
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
      Left            =   7080
      TabIndex        =   13
      Tag             =   "20-Accommodation"
      Top             =   3840
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
   Begin INFOHR_Controls.CodeLookup clpOherCur 
      DataField       =   "ES_OTCUR"
      Height          =   285
      Left            =   8220
      TabIndex        =   17
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpEmployerCur 
      DataField       =   "ES_EMPLOYCUR"
      Height          =   285
      Left            =   8220
      TabIndex        =   18
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpAcomCur 
      DataField       =   "ES_ACOMCUR"
      Height          =   285
      Left            =   8220
      TabIndex        =   19
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin INFOHR_Controls.CodeLookup clpTotCur 
      DataField       =   "ES_TOTCUR"
      Height          =   285
      Left            =   8220
      TabIndex        =   21
      Top             =   4560
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
      Left            =   7080
      TabIndex        =   14
      Tag             =   "20-Accommodation"
      Top             =   4200
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
      Left            =   8220
      TabIndex        =   20
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECUR"
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6840
         TabIndex        =   53
         Top             =   135
         Width           =   1095
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   2850
         TabIndex        =   43
         Top             =   135
         Width           =   585
      End
      Begin VB.Label lblEENum 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   42
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   160
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   44
      Top             =   7365
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin VB.CommandButton cmdRetest 
         Appearance      =   0  'Flat
         Caption         =   "&Continuing Education"
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
         Left            =   120
         TabIndex        =   49
         Tag             =   "Load Beneficiary screen"
         Top             =   120
         Width           =   2085
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7140
         Top             =   165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\newihr\rgedsem.rpt"
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   9360
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   3
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin INFOHR_Controls.DateLookup dlpDatComp 
      DataField       =   "ES_DATCOMP"
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Tag             =   "41-Date when course was completed"
      Top             =   3120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDatRetest 
      DataField       =   "ES_DATRETEST"
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Tag             =   "41-Retest Date"
      Top             =   3480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSMask.MaskEdBox medSCORE 
      DataField       =   "ES_SCORE"
      Height          =   285
      Left            =   1935
      TabIndex        =   5
      Tag             =   "10-Percentage paid by employee"
      Top             =   3840
      Width           =   870
      _ExtentX        =   1535
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
   Begin MSMask.MaskEdBox medContTotal 
      Height          =   285
      Left            =   7080
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4560
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
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2280
      Picture         =   "feesemretest.frx":77D4
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Course Lookup"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   99
      Left            =   120
      TabIndex        =   48
      Top             =   2040
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Retest"
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
      Left            =   120
      TabIndex        =   46
      Top             =   3480
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score (In %)"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Completed"
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
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "ES_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3480
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "ES_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4290
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Type"
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Conducted By      "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Results  "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Hours"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   5580
      TabIndex        =   30
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   Employee $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   5580
      TabIndex        =   29
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Expenses $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   5280
      TabIndex        =   28
      Top             =   3120
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "     Employer $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   5640
      TabIndex        =   27
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   6000
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Accommodation $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   5460
      TabIndex        =   25
      Top             =   3840
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code"
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
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Currency"
      Height          =   255
      Index           =   26
      Left            =   8220
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Learning Material $"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   5460
      TabIndex        =   22
      Top             =   4200
      Width           =   1515
   End
End
Attribute VB_Name = "frmESEMRETEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew  As Integer
Dim fglHredsem As String       '
Dim fglCursName As String      '
Dim fglExtName As String       '
Dim rsDATA As New ADODB.Recordset

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "EdSem_Retest"
    If fglbNew Then
        glbDocKey = 0
    Else
        glbDocKey = rsDATA("ES_ID")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmESEMRETEST")
End Sub

Private Sub cmdRetest_Click()
Unload Me
Load frmESEMINARS
End Sub

Private Sub dlpDatComp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpDatRetest_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpRenewal_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMESEMRETEST"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMESEMRETEST"
End Sub

Private Sub Form_Load()
Dim x%
    
    Screen.MousePointer = DEFAULT
     
    glbOnTop = "FRMESEMRETEST"
    
    If glbtermopen Then
        Data1.ConnectionString = glbAdoIHRAUDIT
    Else
        Data1.ConnectionString = glbAdoIHRDB
    End If
    
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If

    If EERetrieve() = False Then
        MsgBox "Sorry, Employee can not be found"
        If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
    Else
        Me.Show
        If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    End If
    
    
    Screen.MousePointer = HOURGLASS
    
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmESEMRETEST.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
    Call Display_Value
    Call ST_UPD_MODE(True)

    Call INI_Controls(Me)
    For x% = 1 To 15
        Call setCaption(lblTitle(x%))
    Next
    For x% = 0 To 12
        vbxTrueGrid.Columns(x%).Caption = lStr((vbxTrueGrid.Columns(x%).Caption))
    Next

    If glbLinamar Then
        lblTitle(12).FontBold = True
    End If

    Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub imgIcon_Click()
frmESEMList.Show 1
If glbCrsCodeStrArr(17) = "*" Then
    clpCode(0).Text = glbCrsCodeStrArr(2)
    clpCode(1).Text = glbCrsCodeStrArr(1)
    dlpDatComp.Text = glbCrsCodeStrArr(3)
    glbCrsCodeStrArr(17) = ""
End If
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmESEMRETEST")
    Call FillMemoFile(SQLQ, "EdSem_Retest")
End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmESEMRETEST.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub medEECont_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPPE_Change()

End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEECont_LostFocus(Index As Integer)
Call UpConttotal
End Sub

Private Sub medSCORE_GotFocus()
medSCORE = Val(medSCORE) * 100
End Sub

Private Sub medSCORE_LostFocus()
If Len(medSCORE) > 0 Then
    If IsNumeric(medSCORE) Then
        medSCORE = Val(medSCORE) / 100
    End If
End If
End Sub

Private Sub txtCourseHRS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If glbtermopen Then
    SQLQ = "Select * from Term_HREDSEM_RETEST"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE, ES_DATCOMP DESC, ES_EMPNBR"
Else
    SQLQ = "Select * from HREDSEM_RETEST"
    SQLQ = SQLQ & " where ES_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE, ES_DATCOMP DESC, ES_EMPNBR"
End If


Data1.RecordSource = SQLQ
Data1.Refresh

Call UpConttotal

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HREDSEM_RETEST", "SELECT")
Call RollBack '23July99 js

Exit Function

End Function

Private Sub UpConttotal()
Dim x%, xTotal
xTotal = ""
For x% = 0 To 4 '3
    If IsNumeric(medEECont(x%)) Then xTotal = Val(xTotal) + Val(medEECont(x%))
Next
medContTotal = xTotal

End Sub

Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select * from Term_HREDSEM_RETEST"
        SQLQ = SQLQ & " WHERE ES_ID = " & Data1.Recordset!ES_ID
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE, ES_DATCOMP DESC, ES_EMPNBR"
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HREDSEM_RETEST "
        SQLQ = SQLQ & " where ES_ID = " & Data1.Recordset!ES_ID
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE, ES_DATCOMP DESC, ES_EMPNBR"
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

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

imgIcon.Enabled = TF
medContTotal.Enabled = TF
medEECont(0).Enabled = TF
medEECont(1).Enabled = TF
medEECont(2).Enabled = TF
medEECont(3).Enabled = TF
medEECont(4).Enabled = TF
'clpCode(0).Enabled = TF
'clpCode(1).Enabled = TF
'dlpDatComp.Enabled = TF
dlpDatRetest.Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
txtCourseHRS.Enabled = TF
medSCORE.Enabled = TF
dlpRenewal.Enabled = TF
dlpDatRetest.Enabled = TF



glbDocName = "EdSem_Retest"
If gsAttachment_DB Then
    frmDocImport.Visible = True
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("ES_DOCKEY")) Then
                glbDocKey = rsDATA("ES_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("ES_DOCKEY")) Then
                glbDocKey = Data1.Recordset("ES_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
    End If

    Call DispimgIcon(Me, "frmESEMRETEST")
    If gSec_Upd_Education_Seminars And Not glbtermopen Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

'End If

End Sub

Sub cmdCancel_Click()
Dim I%
Dim x
On Error GoTo Can_Err

fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

'If fglbNew Then
'    clpCode(4).Text = ""
'    clpCode(5).Text = ""
'    txtAttHrs.Text = ""
'    txtSkillsExp.Text = ""
'    chkSeniority.Value = False
'    chkIncentive.Value = False
'End If


Call UpConttotal

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREDSEM_RETEST", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESEMRETEST" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


fglHredsem = dlpRenewal.Text
If fglHredsem <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_EDSEM_RETEST where ES_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and ES_DOCKEY=" & glbDocKey & " " '
    End If
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EDSEM_RETEST where ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR = " & glbLEE_ID & " and ES_DOCKEY=" & glbDocKey & " "
    End If
    Data1.Refresh
End If
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If
Call UpConttotal
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREDSEM", "Delete")
Call RollBack '23July99 js

End Sub

Sub cmdNew_Click()
Dim x%
fglbNew = True

'Call ST_UPD_MODE(True)
On Error GoTo AddN_Err

Call SET_UP_MODE

If gsAttachment_DB Then
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
End If

Call Set_Control("B", Me)
rsDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"
For x% = 0 To 4 '3
    medEECont(x%) = 0
Next
Call UpConttotal
'clpCode(1).Enabled = True
dlpDatRetest.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREDSEM_RETEST", "Add")
Call RollBack '23July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
'Dim xChange1, xChange2
Dim x
Dim xID As Long

On Error GoTo Add_Err


If Not chkSemnr() Then Exit Sub

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("ES_ID")
    Data1.Refresh
  Else
    fglHredsem = dlpRenewal.Text
    If fglHredsem <> "" Then
        If Not updFollow("U") Then
                Exit Sub
        End If
    End If
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("ES_ID")
    Data1.Refresh
End If

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
        End If
    End If
    glbDocImpFile = ""
End If



fglbNew = False
Call SET_UP_MODE

Exit Sub
Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREDSEM_RETEST", "Update")
Call RollBack '23July99 js

End Sub

Private Function chkSemnr()
Dim oCode As String, OCodeD As String

chkSemnr = False

If Len(clpCode(1).Text) = 0 Then
    MsgBox lStr("Course Type is a required field")
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox lStr("Course Type code must be valid")
    clpCode(1).SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) > 0 Then
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox lStr("Course Code must be valid")
        clpCode(0).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(2).Text) > 0 Then
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox lStr("Conducted By code must be valid")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox lStr("Results code must be valid")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(0)) < 1 Then
    medEECont(0) = 0
Else
    If Not IsNumeric(medEECont(0)) Then
        MsgBox "Employee's Contribution must be numeric"
        medEECont(0).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(1)) < 1 Then
    medEECont(1) = 0
Else
    If Not IsNumeric(medEECont(1)) Then
        MsgBox "Employer's Contribution must be numeric"
        medEECont(1).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(2)) < 1 Then
    medEECont(2) = 0
Else
    If Not IsNumeric(medEECont(2)) Then
        MsgBox lStr("Other Expenses must be numeric")
        medEECont(2).SetFocus
        Exit Function
    End If
End If
If Len(medEECont(3)) < 1 Then
    medEECont(3) = 0
Else
    If Not IsNumeric(medEECont(3)) Then
        MsgBox lStr("Accommodation must be numeric")
        medEECont(3).SetFocus
        Exit Function
    End If
End If

If Len(txtCourseHRS) > 0 Then
    If Not IsNumeric(txtCourseHRS) Then
        MsgBox lStr("Course Hours must be numeric")
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If


If Len(dlpDatComp.Text) = 0 Then
    If Not IsDate(dlpDatComp.Text) Then
        MsgBox lStr("Date Completed is invalid")
        dlpDatComp.SetFocus
        Exit Function
    End If
End If
If Len(dlpDatRetest.Text) = 0 Then
    If Not IsDate(dlpDatRetest.Text) Then
        MsgBox lStr("Date of Retest is invalid")
        dlpDatRetest.SetFocus
        Exit Function
    End If
End If

If Len(dlpRenewal.Text) > 0 Then
    If Not IsDate(dlpRenewal.Text) Then
        MsgBox lStr("Renewal date is invalid")
        dlpRenewal.SetFocus
        Exit Function
    End If
End If
If glbLinamar Then
    If Val(txtCourseHRS) = 0 Then
        MsgBox lStr("Course Hours is requried field")
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If
If Len(clpEmpCur.Text) > 0 Then
    If clpEmpCur.Caption = "Unassigned" Then
        MsgBox "Employee Currency code must be valid"
        clpEmpCur.SetFocus
        Exit Function
    End If
End If
If Len(clpOherCur.Text) > 0 Then
    If clpOherCur.Caption = "Unassigned" Then
        MsgBox "Other Expenses Currency code must be valid"
        clpOherCur.SetFocus
        Exit Function
    End If
End If
If Len(clpEmployerCur.Text) > 0 Then
    If clpEmployerCur.Caption = "Unassigned" Then
        MsgBox "Employer Currency code must be valid"
        clpEmployerCur.SetFocus
        Exit Function
    End If
End If
If Len(clpAcomCur.Text) > 0 Then
    If clpAcomCur.Caption = "Unassigned" Then
        MsgBox "Accomodations Currency code must be valid"
        clpAcomCur.SetFocus
        Exit Function
    End If
End If
If Len(clpTotCur.Text) > 0 Then
    If clpTotCur.Caption = "Unassigned" Then
        MsgBox "Total Currency code must be valid"
        clpTotCur.SetFocus
        Exit Function
    End If
End If

If ChkDup() Then
        Exit Function
End If

chkSemnr = True

End Function

Private Function ChkDup()
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim Response%
Dim rsTB As New ADODB.Recordset
Dim Title$, Msg1$, DgDef
Dim xESID As Integer

ChkDup = False

'xESID = rsDATA("ES_ID")
Logx = False
If glbtermopen Then
    SQLQ = "SELECT ES_EMPNBR FROM Term_HREDSEM_RETEST WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT ES_EMPNBR FROM HREDSEM_RETEST WHERE ES_EMPNBR = " & glbLEE_ID
End If
If clpCode(0) <> "" Then
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & clpCode(0) & "'"
End If
If clpCode(1) <> "" Then
    SQLQ = SQLQ & " AND ES_CTYPE = '" & clpCode(1) & "'"
End If
If Len(dlpDatComp.Text) > 0 Then
    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(dlpDatComp.Text) & " "
End If
If Len(dlpDatRetest.Text) > 0 Then
    SQLQ = SQLQ & " AND ES_DATRETEST = " & Date_SQL(dlpDatRetest.Text) & " "
End If
'If Not fglbNew Then
'     SQLQ = SQLQ & " AND ES_ID<>" & Data1.Recordset("ES_ID")
'End If
If rsDATA.EditMode <> adEditAdd Then SQLQ = SQLQ & " AND ES_ID<>" & Data1.Recordset("ES_ID")
If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
End If
If Not rsTB.EOF Then Logx = True
rsTB.Close
If Logx = True Then
    Msg$ = "'Course Type' + 'Course Code' + 'Date Completed' + 'Date of Retest' Duplicate!  "
    MsgBox Msg$
    ChkDup = True
End If
End Function

Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String, Edat As String
Dim iRec As Integer
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim Edit1 As Integer
Dim xCourseName As String
newline = Chr$(13) & Chr$(10)
updFollow = False


On Error GoTo CrFollow_Err

If Not IsDate(fglHredsem) Then
    Exit Function
End If

SQLQ = "SELECT ES_COURSE FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & "AND ES_CTYPE = '" & clpCode(1).Text & "' "
SQLQ = SQLQ & "AND ES_CRSCODE = '" & clpCode(0).Text & "' "
If IsDate(dlpDatComp.Text) Then
SQLQ = SQLQ & "AND ES_DATCOMP = " & Date_SQL(dlpDatComp.Text) & " "
End If
rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
xCourseName = "Retest "
If Not rsTemp.EOF Then
    xCourseName = xCourseName & rsTemp("ES_COURSE")
End If
rsTemp.Close

newline = Chr$(13) & Chr$(10)

If IsDate(fglHredsem) Then
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & Val(glbLEE_ID)
    SQLQ = SQLQ & " AND EF_FREAS = 'EDUC'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)
    SQLQ = SQLQ & " AND EF_COMMENTS ='" & xCourseName & "'"

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
    If fglbNew And IsDate(dlpRenewal.Text) Then 'Jaddy 11/15
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpRenewal.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "EDUC"
        rsTB("EF_COMMENTS") = xCourseName
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And IsDate(dlpRenewal.Text) Then 'Jaddy 11/15
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpRenewal.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "EDUC"
        rsTB("EF_COMMENTS") = xCourseName
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And IsDate(dlpRenewal.Text) Then  'Jaddy 11/15 ' edited record
        'dynHRAT.MoveFirst
        'Do Until dynHRAT.EOF
        '    'dynHRAT.Edit
        '    dynHRAT("EF_COMPNO") = "001"
        '    dynHRAT("EF_EMPNBR") = glbLEE_ID
        '    dynHRAT("EF_FDATE") = CVDate(dlpRenewal.Text)
        '    dynHRAT("EF_FREAS") = "EDUC"
        '    dynHRAT("EF_COMMENTS") = xCourseName
        '    dynHRAT("EF_LDATE") = Date
        '    dynHRAT("EF_LTIME") = Time$
        '    dynHRAT("EF_LUSER") = glbUserID
        '    dynHRAT.Update
        '    dynHRAT.MoveNext
        'Loop
        dynHRAT.Close

        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And (Not IsDate(dlpRenewal.Text)) Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Msg = "A record has been deleted from the Follow Up table"
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

        Exit Function
    Else
        updFollow = True
    End If
End If
    
If Not IsDate(dlpRenewal.Text) Then
    updFollow = True
End If

Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

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
UpdateRight = gSec_Upd_Education_Seminars
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

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
Call UpConttotal
End Sub
