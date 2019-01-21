VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEWFCProm 
   Caption         =   "Promotion/Lateral Move"
   ClientHeight    =   8220
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSal 
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
      Height          =   5055
      Left            =   5760
      TabIndex        =   48
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtMarketLine 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "SH_MarketLine"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   55
         Top             =   2640
         Visible         =   0   'False
         Width           =   850
      End
      Begin VB.ComboBox cmbMarketLine 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "00-Market Line"
         Top             =   2430
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtFiscalYear 
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   20
         Tag             =   "00-Fiscal Year"
         Top             =   2100
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox comPayPer 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Tag             =   "01-Choose annum or hour"
         Top             =   750
         Width           =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "SH_EDATE"
         Height          =   285
         Index           =   0
         Left            =   1725
         TabIndex        =   17
         Tag             =   "41-Effective date of salary change"
         Top             =   1080
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medsalary 
         DataField       =   "SH_SALARY"
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Tag             =   "21-Enter salary"
         Top             =   420
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "SH_SREAS1"
         Height          =   285
         Index           =   1
         Left            =   1725
         TabIndex        =   14
         Tag             =   "01-Reason code "
         Top             =   90
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   1
         Left            =   1725
         TabIndex        =   18
         Tag             =   "Next Review Date"
         Top             =   1410
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1725
         TabIndex        =   19
         Tag             =   "00-Enter pay period code"
         Top             =   1755
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDPP"
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SalCode"
         DataField       =   "SH_SALCD"
         DataSource      =   "Data1"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3480
         TabIndex        =   62
         Top             =   840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblMLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
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
         Left            =   3360
         TabIndex        =   59
         Top             =   2430
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblMarketLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   58
         Top             =   2475
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   57
         Top             =   2145
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblBand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Band"
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
         Left            =   3420
         TabIndex        =   56
         Top             =   2100
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Review Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   1455
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   0
         TabIndex        =   53
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label lblReason 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Change"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   52
         Top             =   135
         Width           =   1875
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   51
         Top             =   465
         Width           =   1380
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Per"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   50
         Top             =   795
         Width           =   300
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   49
         Top             =   1125
         Width           =   1245
      End
   End
   Begin VB.Frame fraPos 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   5355
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU3"
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
         Left            =   2520
         TabIndex        =   34
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1410
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU2"
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
         Left            =   2520
         TabIndex        =   33
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU"
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
         Left            =   2520
         TabIndex        =   32
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   750
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         DataField       =   "JH_COMMENT"
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
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "00-Position Comments"
         Top             =   3750
         Width           =   2895
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "JH_SHIFT"
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
         Left            =   2145
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "00-Code assigned to the shift"
         Top             =   3090
         Width           =   810
      End
      Begin VB.TextBox txtComments2 
         Appearance      =   0  'Flat
         DataField       =   "JH_COMMENT2"
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
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "00-Position Notes"
         Top             =   4080
         Width           =   2895
      End
      Begin VB.ComboBox cboShift 
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
         Left            =   3240
         TabIndex        =   10
         Top             =   3090
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "JH_REPTAU4"
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
         Left            =   2520
         TabIndex        =   31
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1755
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmSamuelProfitSharing 
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
         Height          =   285
         Left            =   8880
         TabIndex        =   30
         Top             =   6960
         Visible         =   0   'False
         Width           =   2325
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         DataField       =   "JH_SDATE"
         Height          =   285
         Left            =   1830
         TabIndex        =   1
         Tag             =   "41-Enter Position Start Date"
         Top             =   420
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   12600
         Top             =   4110
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
         Caption         =   "Adodc2"
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
         Left            =   12600
         Top             =   3750
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_DHRS"
         Height          =   285
         Index           =   0
         Left            =   2145
         TabIndex        =   6
         Tag             =   "10-Usual working hours per day"
         Top             =   2100
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_WHRS"
         Height          =   285
         Index           =   1
         Left            =   2145
         TabIndex        =   7
         Tag             =   "10- Number of hours in work week"
         Top             =   2430
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "JH_PHRS"
         Height          =   285
         Index           =   2
         Left            =   2145
         TabIndex        =   8
         Tag             =   "10-Usual working hours per pay period"
         Top             =   2760
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   9
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   12000
         Top             =   3960
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "JH_JREASON"
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   11
         Tag             =   "01-Reason for change in position - Code"
         Top             =   3420
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   2
         Left            =   1830
         TabIndex        =   4
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1410
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   3
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   2
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   750
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         DataField       =   "JH_JOB"
         Height          =   285
         Left            =   1830
         TabIndex        =   0
         Tag             =   "01-Position code"
         Top             =   90
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   5
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1755
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 3"
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
         Left            =   60
         TabIndex        =   47
         Top             =   1455
         Width           =   1290
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 2"
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
         Left            =   60
         TabIndex        =   46
         Top             =   1125
         Width           =   1290
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 1"
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
         Left            =   60
         TabIndex        =   45
         Top             =   3765
         Width           =   555
      End
      Begin VB.Label lblEEStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Change"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   3465
         Width           =   1650
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   60
         TabIndex        =   43
         Top             =   3135
         Width           =   1725
      End
      Begin VB.Label lblHrsPayPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   42
         Top             =   2805
         Width           =   1515
      End
      Begin VB.Label lblHrsWeek 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Week"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   40
         Top             =   2145
         Width           =   930
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   39
         Top             =   795
         Width           =   1515
      End
      Begin VB.Label lblStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   465
         Width           =   885
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   37
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lblComment2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes 2"
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
         Left            =   60
         TabIndex        =   36
         Top             =   4090
         Width           =   555
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 4"
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
         Left            =   60
         TabIndex        =   35
         Top             =   1800
         Width           =   1515
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   255
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
         Left            =   1200
         TabIndex        =   25
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "lblEEName"
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
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   9720
         TabIndex        =   23
         Top             =   4080
         Width           =   945
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   26
      Top             =   7560
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   1335
         TabIndex        =   27
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "New Salary Information"
      Height          =   375
      Left            =   5730
      TabIndex        =   61
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label lblPos 
      Caption         =   "New Position Information"
      Height          =   375
      Left            =   180
      TabIndex        =   60
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmEWFCProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbBAND
Dim fglbSection
Dim lblsalstate0, lblsalstate1, lblsalstate2
Dim xPayrollID
Dim OSalary
Dim xlocUnion
Dim xWFCPosChgEmailBody 'Ticket #29343 Franks 11/01/2016
Dim xIsWFCPosChgEmail As Boolean  'Ticket #29343 Franks 11/01/2016

Private Sub clpJob_KeyUp(KeyCode As Integer, Shift As Integer)
Call WFC_Band
Call Set_MarketLine_List
End Sub

Private Sub clpJob_LostFocus()
Dim xStr, xTmp

Call WFC_Band
Call Set_MarketLine_List

Call WFCReptDisp 'Ticket #29343 Franks 11/01/2016

End Sub

Private Sub WFCReptDisp()
Dim xStr, xTmp
If glbWFC Then 'Ticket #29343 Franks 11/01/2016
    If Len(clpJob.Text) > 0 Then
        xStr = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
        If Len(elpReptAuthShow(0).Text) = 0 And Len(clpJob.Text) > 0 Then
            elpReptAuthShow(0).Text = xStr
        End If
        If Len(xStr) = 0 Then
            If (xlocUnion = "NONE" Or xlocUnion = "EXEC") Then 'Salary employee only
                'lblWFCNote.Visible = True
                xTmp = "No teammate was found for the Position's Reporting Authority #1. Please enter an interim Reporting Authority #1 or contact Total Rewards to update the Position Master"
                MsgBox xTmp
                'Ticket #30180 Franks 05/18/2017 - the following line caused an error, just commented out
                'elpReptAuthShow(0).SetFocus
            End If
        Else
            'lblWFCNote.Visible = False
        End If
    End If
End If

End Sub

Private Sub cmbMarketLine_LostFocus()
txtMarketLine = cmbMarketLine
Call WFC_Band 'Ticket #21677 Franks 03/07/2012
Call Set_MarketLine_List
End Sub

Private Sub cmdClose_Click()
    glbOnTop = ""
    Unload Me
End Sub

Public Sub cmdOK_Click()
Dim Msg$, Title$, DgDef As Variant, Response%
Dim EID&, SEQID&, TermDate$, x%
Dim rsEmp As New ADODB.Recordset
Dim rsPos As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim SQLQ As String
Dim xNewPos As Boolean
Dim xNewSal As Boolean
Dim rJobID
Dim xOldRate, xNewRate

If Not ChkInput() Then Exit Sub

Msg$ = Msg$ & Chr(10) & "Are you sure you want to change Position and Salary information as entered? " 'for this employee ?"

Title$ = "Confirm" ' "Transfer In Employee"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

'MsgBox "This function is not finished yet"
'Exit Sub

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
If rsEmp.State <> 0 Then rsEmp.Close
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsEmp.EOF Then
    Exit Sub
End If

'----------- Position Begin -------------------------
xNewPos = False
SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & "AND JH_JOB = '" & clpJob.Text & "' "
SQLQ = SQLQ & "AND JH_SDATE = " & Date_SQL(dlpStartDate.Text) & " "
If rsPos.State <> 0 Then rsPos.Close
rsPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsPos.EOF Then
    xNewPos = True
    rsPos.AddNew
    rsPos("JH_EMPNBR") = glbLEE_ID
    rsPos("JH_JOB") = clpJob.Text
    rsPos("JH_CURRENT") = -1
End If
rsPos("JH_SDATE") = CVDate(dlpStartDate.Text)
rsPos("JH_JREASON") = clpCode(0).Text
If Len(elpReptAuthShow(0).Text) = 0 Then
    rsPos("JH_REPTAU") = Null
ElseIf IsNumeric(elpReptAuthShow(0).Text) Then
    rsPos("JH_REPTAU") = elpReptAuthShow(0).Text
    If IsDate(dlpStartDate.Text) Then 'Ticket #29343 Franks 11/01/2016
        rsPos("JH_EDATEREPT1") = CVDate(dlpStartDate.Text)
    End If
End If
If Len(elpReptAuthShow(1).Text) = 0 Then rsPos("JH_REPTAU2") = Null Else If IsNumeric(elpReptAuthShow(1).Text) Then rsPos("JH_REPTAU2") = elpReptAuthShow(1).Text
If Len(elpReptAuthShow(2).Text) = 0 Then rsPos("JH_REPTAU3") = Null Else If IsNumeric(elpReptAuthShow(2).Text) Then rsPos("JH_REPTAU3") = elpReptAuthShow(2).Text
If Len(elpReptAuthShow(3).Text) = 0 Then rsPos("JH_REPTAU4") = Null Else If IsNumeric(elpReptAuthShow(3).Text) Then rsPos("JH_REPTAU4") = elpReptAuthShow(3).Text
If IsNumeric(medHours(0).Text) Then rsPos("JH_DHRS") = medHours(0).Text Else rsPos("JH_DHRS") = 0
If IsNumeric(medHours(1).Text) Then rsPos("JH_WHRS") = medHours(1).Text Else rsPos("JH_WHRS") = 0
If IsNumeric(medHours(2).Text) Then rsPos("JH_PHRS") = medHours(2).Text Else rsPos("JH_PHRS") = 0
If Len(txtShift.Text) > 0 Then rsPos("JH_SHIFT") = txtShift.Text
If Len(txtComment.Text) > 0 Then rsPos("JH_COMMENT") = Left(txtComment.Text, 50)
If Len(txtComments2.Text) > 0 Then rsPos("JH_COMMENT2") = Left(txtComments2.Text, 50)
rsPos("JH_LDATE") = Date
rsPos("JH_LTIME") = Time$
rsPos("JH_LUSER") = glbUserID
rsPos.Update
rJobID = rsPos("JH_ID")

Call AUDITPSTN("A")

rsPos.Close

If xNewPos Then
    Call Set_Current_Pos_Flag(glbLEE_ID)
    Call WFCPosSkillsUpd(glbLEE_ID, clpJob.Text, dlpStartDate.Text)
    'Call Employee_Master_Integration(glbLEE_ID) 'for AT
End If

'----------- Position end -------------------------

'=========== Salary Begin ==============================
xNewSal = False
SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & "AND SH_JOB = '" & clpJob.Text & "' "
SQLQ = SQLQ & "AND SH_EDATE = " & Date_SQL(dlpDate(0).Text) & " "
If rsSal.State <> 0 Then rsSal.Close
rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
xOldRate = 0
If rsSal.EOF Then
    xNewSal = True
    xOldRate = getCurWFCSalary(glbLEE_ID)
    rsSal.AddNew
    rsSal("SH_EMPNBR") = glbLEE_ID
    rsSal("SH_JOB") = clpJob.Text
    rsSal("SH_CURRENT") = -1
    rsSal("SH_JOB_ID") = rJobID
    If xOldRate > 0 Then
        If IsNumeric(medsalary.Text) Then xNewRate = medsalary.Text Else xNewRate = 0
        rsSal("SH_SALCHG1") = xNewRate - xOldRate
        If xOldRate = 0 Then rsSal("SH_SALPC1") = 1 Else rsSal("SH_SALPC1") = rsSal("SH_SALCHG1") / xOldRate
    End If
End If
OSalary = xOldRate
rsSal("SH_SDATE") = CVDate(dlpStartDate.Text)
rsSal("SH_SREAS1") = clpCode(1).Text
If IsNumeric(medsalary.Text) Then rsSal("SH_SALARY") = medsalary.Text Else rsSal("SH_SALARY") = 0
rsSal("SH_SALCD") = Left(lblSalCode.Caption, 1)
If IsDate(dlpDate(0).Text) Then rsSal("SH_EDATE") = CVDate(dlpDate(0).Text)
If IsDate(dlpDate(1).Text) Then rsSal("SH_NEXTDAT") = CVDate(dlpDate(1).Text)
If Len(clpCode(2).Text) > 0 Then rsSal("SH_PAYP") = clpCode(2).Text
If IsNumeric(medHours(1).Text) Then rsSal("SH_WHRS") = medHours(1).Text Else rsSal("SH_WHRS") = 0

If txtFiscalYear.Visible And Len(txtFiscalYear.Text) > 0 Then rsSal("SH_FISCALYEAR") = txtFiscalYear.Text
If cmbMarketLine.Visible And Len(cmbMarketLine.Text) > 0 Then rsSal("SH_MARKETLINE") = cmbMarketLine.Text
If lblBand.Visible And Len(fglbBAND) > 0 Then rsSal("SH_BAND") = fglbBAND

rsSal("SH_GRADE") = "00"
If txtFiscalYear.Visible Then
    rsSal("SH_COMPA") = Get_WFC_COMPA
    If IsNumeric(lblsalstate0) Then rsSal("SH_LDOLLARS") = lblsalstate0
    If IsNumeric(lblsalstate1) Then rsSal("SH_MDOLLARS") = lblsalstate1
    If IsNumeric(lblsalstate1) Then rsSal("SH_HDOLLARS") = lblsalstate1
Else
    rsSal("SH_COMPA") = 0
End If
rsSal("SH_SECTION") = fglbSection
rsSal("SH_PAYROLL_ID") = xPayrollID
rsSal("SH_TRANSDATE") = Date '
rsSal("SH_LDATE") = Date '
rsSal("SH_LTIME") = Time$
rsSal("SH_LUSER") = glbUserID
rsSal.Update

Call AUDITSALY("A")

Call AUDIT_NGS_TRANS
If rsSal.State <> 0 Then rsSal.Close

If xNewSal Then
    Call Set_Current_Sal_Flag(glbLEE_ID)
    Call updBenefitForSalDEPN(glbLEE_ID)
End If
'=========== Salary end ==============================

If glbCandidate > 0 Then
    rsEmp("ED_CANDIDATE") = glbCandidate
    rsEmp.Update
    rsEmp.Close
    Call WFCHRSoftProcUpt("frmEWFCProm")
End If
    
If xIsWFCPosChgEmail Then  'Ticket #29343 Franks 11/01/2016
    If gsEMAIL_ONPOSITION Then
        If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
            Call WFCPubPosChangedcmdEmail(glbLEE_ID, xWFCPosChgEmailBody, "info:HR Position Reporting Authority Change Notice")
        End If
    End If
End If
    
If xNewPos Or xNewSal Then
    Call Employee_Master_Integration(glbLEE_ID) 'for AT
End If

MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT
Unload Me

End Sub

Private Sub comPayPer_LostFocus()
If comPayPer.ListIndex = 0 Then
    lblSalCode.Caption = "A"
ElseIf comPayPer.ListIndex = 1 Then
    lblSalCode.Caption = "H"
ElseIf comPayPer.ListIndex = 2 Then 'Ticket #14645
    lblSalCode.Caption = "M"
ElseIf comPayPer.ListIndex = 3 Then
    lblSalCode.Caption = "D"
End If
End Sub

Private Sub dlpDate_Change(Index As Integer)
    If (xlocUnion = "NONE" Or xlocUnion = "EXEC") Then
        If Index = 0 Then
            dlpDate(1).Text = getNextReviewDate(dlpDate(0).Text, dlpDate(1).Text)
        End If
    End If
End Sub

Private Sub Form_Activate()
glbOnTop = "frmEWFCProm"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmEWFCProm"
End Sub

Private Sub Form_Load()
glbOnTop = "frmEWFCProm"

lblTitle(28).Caption = lStr("Pay Period") 'Ticket #21988 Franks 05/02/2012

clpJob.TextBoxWidth = 1315

comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "
comPayPer.AddItem "Daily "

Call ShiftItem

Call DecSetup

If Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    If glbHRSoftType = "PROM" Then Me.Caption = "Promotion - " & Left$(glbLEE_SName, 5)
    If glbHRSoftType = "LATM" Then Me.Caption = "Lateral Move - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

lblEENum.Caption = ShowEmpnbr(glbLEE_ID)
        
Call WFCHRSoftDispValues

Call INI_Controls(Me)
End Sub



Public Property Get ChangeAction() As UpdateStateEnum
 ChangeAction = OPENING
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)

End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateTransEmp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Terminations
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = True 'xUpdateable
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

UpdateState = OPENING
TF = True
Call set_Buttons(UpdateState)
If Not UpdateRight Then
    TF = False
    'Call UPDMOD
End If

End Sub

Private Sub WFCHRSoftDispValues()
Dim rsEmp As New ADODB.Recordset
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp
Dim xSalCD

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
If rsEmp.State <> 0 Then rsEmp.Close
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
lblsalstate0 = 0
lblsalstate1 = 0
lblsalstate2 = 0
fglbSection = ""
xPayrollID = ""
If Not rsEmp.EOF Then
    fglbSection = rsEmp("ED_SECTION")
    xPayrollID = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

'xHRSoftUpt = False
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    SQLQ = SQLQ & "AND SF_UPT_PROCESSED = 0 "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xlocUnion = ""
    If Not rsCanid.EOF Then
        'for position  - begin
        If Not IsNull(rsCanid("SF_POSITIONCODE")) Then
            clpJob.Text = rsCanid("SF_POSITIONCODE")
        End If
        If Not IsNull(rsCanid("SF_STARTDATE")) Then dlpStartDate = rsCanid("SF_STARTDATE")
        If Not IsNull(rsCanid("SF_ORG")) Then
            xlocUnion = rsCanid("SF_ORG")
            Call WFCDefaultHours(rsCanid("SF_ORG")) 'hours
        End If
        If glbHRSoftType = "PROM" Then clpCode(0).Text = glbHRSoftType
        If glbHRSoftType = "LATM" Then clpCode(0).Text = glbHRSoftType
        'for position  - end
        
        'for salary  - begin
        If glbHRSoftType = "PROM" Then clpCode(1).Text = glbHRSoftType
        If glbHRSoftType = "LATM" Then clpCode(1).Text = glbHRSoftType
        If Not IsNull(rsCanid("SF_SALARY")) Then medsalary.Text = rsCanid("SF_SALARY")
        xSalCD = ""
        If Not IsNull(rsCanid("SF_SALARYFREQUENCY")) Then
            If rsCanid("SF_SALARYFREQUENCY") = "Annum" Then xSalCD = "A"
            If rsCanid("SF_SALARYFREQUENCY") = "Hour" Then xSalCD = "H"
            If rsCanid("SF_SALARYFREQUENCY") = "Monthly" Then xSalCD = "M"
            If rsCanid("SF_SALARYFREQUENCY") = "Daily" Then xSalCD = "D"
            
            'Ticket #24620 Franks 12/02/2013
            clpCode(2).Text = "W"
            If xSalCD = "A" Then clpCode(2).Text = "SM"
            If xSalCD = "M" Then clpCode(2).Text = "M"
    
        End If
        If Len(xSalCD) > 0 Then
            lblSalCode.Caption = xSalCD
            Call comPayPerList(xSalCD)
        End If
        If Not IsNull(rsCanid("SF_STARTDATE")) Then dlpDate(0).Text = rsCanid("SF_STARTDATE")
        'for salary  - end
        
        Call WFCReptDisp
        
    End If
End If
End Sub

Private Sub comPayPerList(xSalCD)
    comPayPer.ListIndex = -1
    If xSalCD = "A" Then comPayPer.ListIndex = 0
    If xSalCD = "H" Then comPayPer.ListIndex = 1
    If xSalCD = "M" Then comPayPer.ListIndex = 2
    If xSalCD = "D" Then comPayPer.ListIndex = 3
End Sub

Private Sub DecSetup()
If glbCompDecHR = 3 Then
    medsalary.Format = "#,##0.000;(#,##0.000)"
End If
If glbCompDecHR = 4 Then
    medsalary.Format = "#,##0.0000;(#,##0.0000)"
End If
End Sub

Private Sub WFCDefaultHours(xUnion) 'Ticket #24451 Franks 10/17/2013
    If Len(xUnion) > 0 Then
        If (xUnion = "NONE" Or xUnion = "EXEC" Or xUnion = "-NON" Or xUnion = "-EXE") Then  'salaried
            medHours(0).Text = 8
            medHours(1).Text = 40
            medHours(2).Text = 86.67
            clpCode(2).Text = "SM"
            
            txtFiscalYear.Visible = True
            cmbMarketLine.Visible = True
            lblFiscalYear.Visible = True
            lblMarketLine.Visible = True
            lblMLine.Visible = True
        Else 'hourly
            medHours(0).Text = 8
            medHours(1).Text = 40
            medHours(2).Text = 40
            clpCode(2).Text = "W"
            
            txtFiscalYear.Visible = False
            cmbMarketLine.Visible = False
            lblFiscalYear.Visible = False
            lblMarketLine.Visible = False
            lblMLine.Visible = False
        End If
    End If
    
End Sub

Private Sub cboShift_Change()
    'If cboShift.ListIndex > -1 Then
        txtShift.Text = cboShift.Text
    'End If
End Sub

Private Sub cboShift_Click()
    'If cboShift.ListIndex > -1 Then
        txtShift.Text = cboShift.Text
    'End If
End Sub

Private Sub ShiftItem()
    cboShift.Left = txtShift.Left
    cboShift.Visible = True
    txtShift.Visible = False
    cboShift.AddItem "NS"
    cboShift.AddItem "1"
    cboShift.AddItem "2"
    cboShift.AddItem "3"
    cboShift.AddItem "4"
    cboShift.AddItem "5"
    cboShift.AddItem "6"
    cboShift.AddItem "A"
    cboShift.AddItem "B"
    cboShift.AddItem "C"
    cboShift.AddItem "D"
    cboShift.AddItem "E"
    cboShift.AddItem "F"
    cboShift.AddItem "M"
    cboShift.AddItem "Q"
    cboShift.AddItem "R"
    cboShift.AddItem "S"
    cboShift.AddItem "T"
    cboShift.AddItem "W"

End Sub

Private Function ChkInput()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
Dim x%
Dim Msg$, Response%
Dim xDivCountry As String
ChkInput = False

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
If rsEmp.State <> 0 Then rsEmp.Close
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEmp.EOF Then Exit Function

'fglbSection = rsEmp("ED_SECTION")
'xPayrollID = rsEmp("ED_PAYROLL_ID")
'lblsalstate0 = 0
'lblsalstate1 = 0
'lblsalstate2 = 0

'position - begin
If Len(clpJob) = 0 Then
    MsgBox "Position Code is a required field"
    clpJob.SetFocus
    Exit Function
Else
    If clpJob.Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpJob.SetFocus
        Exit Function
    End If
End If

If Len(dlpStartDate) > 0 Then
    If Not IsDate(dlpStartDate) Then
        MsgBox "Invalid Position Start Date"
        dlpStartDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Position Start Date is a required field."
    dlpStartDate.SetFocus
    Exit Function
End If

For x% = 0 To 3 '2
    If elpReptAuthShow(x%) = "0" Then elpReptAuthShow(x%) = ""
    If Len(elpReptAuthShow(x%)) > 0 Then
        If elpReptAuthShow(x%).Caption = "Unassigned" Then
            MsgBox "Rept. Authority Employee # not valid. Check Employee # and re-enter!"
            elpReptAuthShow(x%).SetFocus
            Exit Function
        End If
    End If
Next

If Len(elpReptAuthShow(0).Text) = 0 Then
    MsgBox "Rept. Authority 1 is required."
    elpReptAuthShow(0).SetFocus
    Exit Function
End If


xIsWFCPosChgEmail = False
If (xlocUnion = "NONE" Or xlocUnion = "EXEC") Then  'Salary employee only 'Ticket #29343 Franks 11/01/2016
    If Len(elpReptAuthShow(0).Text) > 0 Then
        If IsRept1PosNotMatchPosMaster(elpReptAuthShow(0).Text, clpJob.Text) Then
            glbMsgCustomVal = 11
            frmMsgDialog.Show 1
            'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
            If glbMsgCustomVal = 2 Then 'If <<Cancel>> is checked, undo the change.
                elpReptAuthShow(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
                Exit Function
            End If
            'If <<Continue>> is checked, send email
            If gsEMAIL_ONPOSITION Then
                xWFCPosChgEmailBody = "This Reporting Authority #1 " & elpReptAuthShow(0).Text & " " & GetEmpData(elpReptAuthShow(0).Text, "ED_SURNAME") & "," & GetEmpData(elpReptAuthShow(0).Text, "ED_FNAME") & " "
                xWFCPosChgEmailBody = xWFCPosChgEmailBody & "is not associated with this position and may cause a break in the organization chain."
                xIsWFCPosChgEmail = True
            End If
        End If
    End If
End If
        
If Not IsNumeric(medHours(0)) Then
    MsgBox "Hours/Day is required"
    medHours(0).SetFocus
    Exit Function
End If
If Not IsNumeric(medHours(1)) Then
    MsgBox "Hours/Week is required"
    medHours(1).SetFocus
    Exit Function
End If
If Not IsNumeric(medHours(2)) Then
    MsgBox "Hours/Per Period is required"
    medHours(2).SetFocus
    Exit Function
End If
For x% = 0 To 2
    If Not IsNumeric(medHours(x%)) Then medHours(x%) = 0
Next

If Len(clpCode(0)) = 0 Then
    MsgBox "Position Reason for Change is a required field"
     clpCode(0).SetFocus
    Exit Function
Else
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(0).SetFocus
        Exit Function
    End If
End If

If glbAdv And Not glbWFCFullRights Then 'Ticket #13867
    If IsNull(rsEmp("ED_BONUSDEPT")) Or Len(rsEmp("ED_BONUSDEPT")) = 0 Then
        If Len(txtShift.Text) = 0 Then
            MsgBox lStr("Shift is a required field")
            txtShift.SetFocus
            Exit Function
        End If
    End If
End If
'Position end ----------------------------

'Salary -begin ---------------------------
If Len(clpCode(1)) = 0 Then
    MsgBox "Salary Reason for Change is a required field"
     clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(1).SetFocus
        Exit Function
    End If
End If
If Len(medsalary) < 1 Then
    MsgBox "Salary is required"
    medsalary.SetFocus
    Exit Function
End If
If Val(medsalary) <= 0 Then
    MsgBox "Salary is required"
    medsalary.SetFocus
    Exit Function
End If
If comPayPer.ListIndex = -1 Or lblSalCode = "" Then
    MsgBox "Per is required field"
    comPayPer.SetFocus
    Exit Function
End If
If Len(dlpDate(0)) > 0 Then
    If Not IsDate(dlpDate(0)) Then
        MsgBox "Invalid Salary Effective Date"
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Salary Effective Date is a required field."
    dlpDate(0).SetFocus
    Exit Function
End If
If Trim(comPayPer.Text) = "Annum" Or Trim(comPayPer.Text) = "Monthly" Then
    If Not IsDate(dlpDate(1).Text) Then
        Msg$ = "Next Review is required"
        dlpDate(1).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
'If Len(dlpDate(1)) > 0 Then
'    If Not IsDate(dlpDate(0)) Then
'        MsgBox "Invalid Date of Next Review Date"
'        dlpDate(1).SetFocus
'        Exit Function
'    End If
'Else
'    MsgBox "Next Review Date is a required field."
'    dlpDate(1).SetFocus
'    Exit Function
'End If

If Len(clpCode(2)) = 0 Then
    MsgBox lStr("Pay Period is a required field")
    clpCode(2).SetFocus
    Exit Function
Else
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Pay Period")
         clpCode(2).SetFocus
        Exit Function
    End If
End If

If txtFiscalYear.Visible Then
    If Len(txtFiscalYear) < 1 Then
        Msg$ = "Fiscal Year is required"
        txtFiscalYear.SetFocus
        MsgBox Msg$
        Exit Function
    Else
        If Not IsNumeric(txtFiscalYear) Then
            Msg$ = "Invalid Fiscal Year"
            txtFiscalYear.SetFocus
            MsgBox Msg$
            Exit Function
        End If
    End If
End If
If cmbMarketLine.Visible And Len(cmbMarketLine.Text) < 1 Then
    Msg$ = "Market Line is required"
    cmbMarketLine.SetFocus
    MsgBox Msg$
    Exit Function
End If

ChkInput = True

End Function

Private Sub txtFiscalYear_LostFocus()
If Len((txtFiscalYear)) > 0 Then
    If Not IsNumeric(txtFiscalYear) Then
        MsgBox "Invalid Fiscal Year."
        txtFiscalYear.SetFocus
    End If
    If Val(txtFiscalYear) < 1900 Or Val(txtFiscalYear) > 3000 Then
        MsgBox "Invalid Fiscal Year."
        txtFiscalYear.SetFocus
    End If
End If
Call WFC_Band 'Ticket #21677 Franks 03/07/2012
Call Set_MarketLine_List
End Sub

Private Sub WFC_Band()
    If glbWFC Then
        If Not clpJob.Caption = "Unassigned" Then
            fglbBAND = getPosBand(clpJob.Text)
            If Len(fglbBAND) > 0 Then
                lblBand.Caption = "Band: " & fglbBAND
            Else
                lblBand.Caption = ""
            End If
            lblBand.Top = lblFiscalYear.Top
            lblBand.Visible = lblMarketLine.Visible
        End If
    End If
End Sub

Sub Set_MarketLine_List()
Dim rsWFC As New ADODB.Recordset
Dim x%, I%
Dim xItemAdd
Dim SQLQ

'SQLQ = "select MarketLine from WFC_Salary_Administration "
SQLQ = "select * from WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & fglbBAND & "'"
If Len(fglbSection) > 0 Then
    SQLQ = SQLQ & " AND SectionCode ='" & fglbSection & "'"
End If
If Len(txtFiscalYear) > 0 Then
    SQLQ = SQLQ & " AND FiscalYear =" & txtFiscalYear & ""
End If
'SQLQ = SQLQ & " group by MarketLine"
SQLQ = SQLQ & " ORDER by MarketLine"

rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset
x% = 0
cmbMarketLine.Clear
Do Until rsWFC.EOF
    cmbMarketLine.AddItem rsWFC("marketline")
    If rsWFC("marketline") = txtMarketLine Then
        cmbMarketLine.ListIndex = x%
        lblsalstate0 = rsWFC("LDOLLARS")
        lblsalstate1 = rsWFC("MDOLLARS")
        lblsalstate2 = rsWFC("HDOLLARS")
    End If
    x% = x% + 1
    rsWFC.MoveNext
Loop
rsWFC.Close
'MarketLine_Desc Me
Call SalMarketLineDesc

End Sub

Private Function getPosBand(xPCODE)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval
    retval = ""
    SQLQ = "SELECT JB_CODE, JB_BAND FROM HRJOB WHERE JB_CODE = '" & xPCODE & "' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("JB_BAND")) Then
            retval = rsTemp("JB_BAND")
        End If
    End If
    getPosBand = retval
End Function

Private Sub SalMarketLineDesc()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    If Len(Trim(cmbMarketLine)) > 0 Then
        SQLQ = "SELECT TB_KEY,TB_DESC FROM HRTABL WHERE TB_NAME ='WFML' AND TB_KEY ='" & cmbMarketLine & "' "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            lblMLine.Caption = rsTemp("TB_DESC")
        End If
        rsTemp.Close
    End If
End Sub

Private Sub Set_Current_Pos_Flag(xEmpNo)
Dim SQLQ As String, Msg$, x
Dim dyn_HRJOBHIS As New ADODB.Recordset

On Error GoTo CurFlgErr
If glbMulti Then Exit Sub

SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) AND  JH_EMPNBR = " & xEmpNo & " ORDER BY JH_SDATE DESC, JH_ID DESC"

dyn_HRJOBHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dyn_HRJOBHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
gdbAdoIhr001.BeginTrans

If dyn_HRJOBHIS.RecordCount > 0 Then dyn_HRJOBHIS.MoveFirst
dyn_HRJOBHIS("JH_CURRENT") = True
'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    dyn_HRJOBHIS("JH_TRK_CRS_RENEWAL") = False
'End If
dyn_HRJOBHIS.Update

dyn_HRJOBHIS.MoveNext

While Not dyn_HRJOBHIS.EOF
    If dyn_HRJOBHIS("JH_CURRENT") <> 0 Then
        dyn_HRJOBHIS("JH_CURRENT") = False
        'Ticket #21511 - If the Current not checked then Default Position should be Off too.
        dyn_HRJOBHIS("JH_POSITION_CONTROL") = "NO"
        dyn_HRJOBHIS.Update
    End If
    dyn_HRJOBHIS.MoveNext
Wend
gdbAdoIhr001.CommitTrans

dyn_HRJOBHIS.Close

Exit Sub

CurFlgErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_JOB_HIS", "Add")
Call RollBack '26July99 js

End Sub

Private Function AUDITPSTN(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xPT, xDiv
Dim HRChangs As New Collection
Dim UpdateAudit As Boolean
Dim UptPositionDate As Date

On Error GoTo AUDIT_ERR

AUDITPSTN = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
Else
    xPT = ""
    xDiv = ""
End If

'Removed * Ticket #9899
Dim strFields As String
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_PHRS, AU_OLDPHRS, AU_WHRS, AU_OLDWHRS, AU_DHRS, AU_OLDDHRS, "
strFields = strFields & "AU_JOB, AU_SJDATE, AU_JREASON, AU_LEADHAND, AU_LABOURCD, AU_LABOUREDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_ORG, AU_BILLINGRATE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If IsNumeric(medHours(2).Text) Then rsTA("AU_PHRS") = medHours(2).Text
If IsNumeric(medHours(1).Text) Then rsTA("AU_WHRS") = medHours(1).Text
If IsNumeric(medHours(0).Text) Then rsTA("AU_DHRS") = medHours(0).Text

rsTA("AU_JOB") = clpJob.Text
rsTA("AU_SJDATE") = dlpStartDate.Text
rsTA("AU_JREASON") = clpCode(0).Text


rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID

If dlpStartDate > Date Then
    rsTA("AU_LDATE") = dlpStartDate
Else
    rsTA("AU_LDATE") = Date
End If


rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX

Dim rsEmp As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA.Update

MODNOUPD:
AUDITPSTN = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js
Resume Next
End Function


Private Function Get_WFC_COMPA()
Dim xDollear
Dim retval

    retval = 0
    If IsNumeric(lblsalstate1) Then xDollear = lblsalstate1 Else xDollear = 0

    If Val(xDollear) <> 0 Then
            retval = (Val(medsalary) / xDollear) * 100
    End If
    If retval > 999.99 Then retval = "999.99"

    Get_WFC_COMPA = retval
End Function

Private Function getCurWFCSalary(xEmpNo)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval
    retval = 0
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & xEmpNo & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = rsTemp("SH_SALARY")
    End If
    rsTemp.Close
    getCurWFCSalary = retval
End Function

Private Function AUDITSALY(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim SQLQ As String, strFields As String
Dim xEffDateUpd, xSalUpd As Boolean

On Error GoTo AUDIT_ERR

AUDITSALY = False

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
'strFields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_GRID, AU_SALARY, AU_OLDSAL, AU_WHRS, AU_SALCD, "
'Added by Bryan 27/09/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'muskoka
    strFields = strFields & "AU_TOTAL, AU_VPREMIUM, AU_VGROUP, AU_VSTEP, "
End If
If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #22952 Franks 12/10/2012
     strFields = strFields & "AU_TOTAL, AU_VPREMIUM, "
End If
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    'North Perth Ticket #19209 Franks 05/18/2011
    'KN&V Ticket #21097 Franks 11/02/2011
    strFields = strFields & "AU_VGROUP, "
End If
strFields = strFields & "AU_JOB, AU_SEDATE, AU_SREASON, AU_PAYP, AU_OLDPAYP, "
strFields = strFields & "AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_JOB "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
'rsTA("AU_GRID") = clpGrid.Text
rsTA("AU_SALARY") = medsalary
If OSalary > 0 Then rsTA("AU_OLDSAL") = OSalary
If IsNumeric(medHours(1).Text) Then rsTA("AU_WHRS") = medHours(1).Text
rsTA("AU_SALCD") = lblSalCode
rsTA("AU_JOB") = clpJob.Text
If IsDate(dlpDate(0).Text) Then rsTA("AU_SEDATE") = dlpDate(0).Text
If Len(clpCode(1).Text) > 0 Then rsTA("AU_SREASON") = clpCode(1).Text
If Len(clpCode(2).Text) > 0 Then rsTA("AU_PAYP") = clpCode(2).Text
If IsDate(dlpDate(1).Text) Then rsTA("AU_SNDATE") = dlpDate(1).Text
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID

If ACTX = "A" Then
        rsTA("AU_LDATE") = dlpDate(0).Text
Else
    If dlpDate(0) > Date Then
        rsTA("AU_LDATE") = dlpDate(0)
    Else
        rsTA("AU_LDATE") = Date
    End If
End If

rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
Dim rsEmp As New ADODB.Recordset
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA.Update
' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
Call Pause(0.5)

MODNOUPD:
AUDITSALY = True

Exit Function
AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '28July99 js
Resume Next

End Function

Private Function AUDIT_NGS_TRANS()
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xDate1, xDate2
Dim xlocFlag As Boolean
Dim xOldVal, xNewVal


On Error GoTo AUDIT_ERR

SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEmpee.EOF Then
    Exit Function
Else
    If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
    If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
End If
rsEmpee.Close

'No NGS Sub Group, skip
If Len(glbWFCNGSSubGroup) = 0 Then Exit Function

xLDate = Date

'NGS field changes --------------------------------------
xlocFlag = False
'If OEDate <> dlpDate(0).Text Then
'    If Len(OEDate) = 0 Then
'        xlocFlag = True
'    Else
'        If Not (CVDate(OEDate) = CVDate(dlpDate(0).Text)) Then
'            xlocFlag = True
'        End If
'    End If
'    If xlocFlag Then
        xDate1 = "" 'OEDate
        xDate2 = dlpDate(0).Text
        Call NGSAuditAdd(glbLEE_ID, "M", "Salary History", "Effective Date", xDate1, xDate2, xLDate)
'    End If
'End If

'Salary amount
xOldVal = OSalary
xNewVal = medsalary
'If Not (OSalary = medsalary) Then
    If xOldVal = 0 Then xOldVal = ""
    Call NGSAuditAdd(glbLEE_ID, "M", "Salary History", "Salary", xOldVal, xNewVal, xLDate)
'End If


AUDIT_NGS_TRANS = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING NGS AUDIT RECORD", "NGS AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
    
End Function

Private Sub Set_Current_Sal_Flag(xEmpNo)
Dim SQLQ As String, Msg$
Dim dyn_HRSALHIS As New ADODB.Recordset

On Error GoTo SCFError
If glbMulti Then Exit Sub

SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & xEmpNo & " "
SQLQ = SQLQ & "ORDER BY SH_EDATE DESC, SH_ID DESC "
dyn_HRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

Screen.MousePointer = HOURGLASS

If dyn_HRSALHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

If dyn_HRSALHIS.RecordCount > 0 Then dyn_HRSALHIS.MoveFirst
dyn_HRSALHIS("SH_CURRENT") = True
dyn_HRSALHIS.Update

Do Until dyn_HRSALHIS.EOF
    dyn_HRSALHIS.MoveNext
    If dyn_HRSALHIS.EOF Then Exit Do
    
    'Hemu - 07/07/2003 Begin - to improve speed, Jaddy suggested
    If dyn_HRSALHIS("SH_CURRENT") <> 0 Then
        dyn_HRSALHIS("SH_CURRENT") = False
        dyn_HRSALHIS.Update
    End If
    'Hemu - 07/07/2003 End
Loop
dyn_HRSALHIS.Close

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Screen.MousePointer = DEFAULT

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

End Sub

Private Function getNextReviewDate(xEDate, xDefDate)
Dim retval
    retval = xDefDate 'dlpDate(1).Text
    If IsDate(xEDate) Then
        retval = CVDate(("May 1, ") & Year(xEDate) + 1)
    End If
    getNextReviewDate = retval
End Function

