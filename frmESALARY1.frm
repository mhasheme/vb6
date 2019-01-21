VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmESALARY11 
   Caption         =   "Salary History"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSalary2 
      Height          =   2295
      Left            =   120
      TabIndex        =   47
      Top             =   5760
      Width           =   11415
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         DataField       =   "SH_COMMENT"
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   55
         Tag             =   "00-Position Comments"
         Top             =   840
         Width           =   3405
      End
      Begin VB.TextBox txtUserSys 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "SH_COMPA_USER"
         Height          =   285
         Left            =   3570
         TabIndex        =   54
         Top             =   1170
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbMarketLine 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Tag             =   "00-Market Line"
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox txtMarketLine 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "SH_MarketLine"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6330
         TabIndex        =   52
         Top             =   810
         Visible         =   0   'False
         Width           =   850
      End
      Begin VB.OptionButton optUserSys 
         Caption         =   "System"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   51
         Top             =   1140
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optUserSys 
         Caption         =   "User"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   50
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox txtPayrollID 
         Appearance      =   0  'Flat
         DataField       =   "SH_PAYROLL_ID"
         Height          =   285
         Left            =   6300
         MaxLength       =   15
         TabIndex        =   49
         Tag             =   "00-Payroll ID"
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtFiscalYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9660
         MaxLength       =   4
         TabIndex        =   48
         Tag             =   "00-Fiscal Year"
         Top             =   435
         Visible         =   0   'False
         Width           =   855
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "SH_EDATE"
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   56
         Tag             =   "41-Effective date of salary change"
         Top             =   420
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "SH_NEXTDAT"
         Height          =   285
         Index           =   1
         Left            =   6300
         TabIndex        =   57
         Tag             =   "40-Next Date to Review Salary"
         Top             =   1530
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "SH_PAYP"
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   58
         Tag             =   "00-Enter pay period code"
         Top             =   1500
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDPP"
      End
      Begin MSMask.MaskEdBox mskCampa 
         Height          =   285
         Left            =   4080
         TabIndex        =   59
         Top             =   1140
         Width           =   1095
         _ExtentX        =   1931
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
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   8325
         TabIndex        =   60
         Tag             =   "00-Section - Code"
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "SH_TRANSDATE"
         Height          =   285
         Index           =   2
         Left            =   6300
         TabIndex        =   61
         Tag             =   "40-Transaction Date"
         Top             =   1890
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   83
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Compa-Ratio"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   82
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblCompaNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "SH_COMPA"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1890
         TabIndex        =   81
         Top             =   1170
         Width           =   90
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   80
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours per Week"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   5280
         TabIndex        =   79
         Top             =   150
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Review"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   5280
         TabIndex        =   78
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblWhrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         DataField       =   "SH_WHRS"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6690
         TabIndex        =   77
         Top             =   150
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lblsalstate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   6300
         TabIndex        =   75
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblsalstate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7260
         TabIndex        =   74
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblsalstate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   8220
         TabIndex        =   73
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblMarketLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   72
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label lblMLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7560
         TabIndex        =   71
         Top             =   840
         Width           =   840
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Scale"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   5280
         TabIndex        =   70
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label lblPayID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   69
         Top             =   450
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8640
         TabIndex        =   68
         Top             =   450
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblPlant 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7440
         TabIndex        =   67
         Top             =   120
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1905
         Width           =   1095
      End
      Begin VB.Label lblUserDesc 
         Caption         =   "lblUserDesc"
         Height          =   255
         Left            =   1440
         TabIndex        =   65
         Top             =   1905
         Width           =   2775
      End
      Begin VB.Label lblLambtonJob 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
         Height          =   195
         Left            =   5280
         TabIndex        =   64
         Top             =   870
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label txtLambtonJob 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6330
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   4755
         TabIndex        =   62
         Top             =   1935
         Width           =   1455
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SH_LUSER"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SH_LTIME"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   330
      MaxLength       =   25
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   8670
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SH_LDATE"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   8670
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtWHRS 
      Appearance      =   0  'Flat
      DataField       =   "SH_WHRS"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4980
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "10-Hours per Week"
      Top             =   2460
      Width           =   495
   End
   Begin Threed.SSFrame fraSalary 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   4350
      Width           =   9045
      _Version        =   65536
      _ExtentX        =   15954
      _ExtentY        =   2672
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin VB.ComboBox comSalScale 
         Height          =   315
         Left            =   6300
         TabIndex        =   6
         Tag             =   "00-Position Grid Steps"
         Top             =   180
         Width           =   675
      End
      Begin VB.ComboBox comPayPer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "01-Choose annum or hour"
         Top             =   180
         Width           =   1215
      End
      Begin VB.ComboBox cboVGRoup 
         Height          =   315
         Left            =   5280
         TabIndex        =   4
         Top             =   735
         Width           =   2055
      End
      Begin VB.ComboBox cboVStep 
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtVGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtVStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "01-Country"
         Top             =   1200
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSMask.MaskEdBox medsalary 
         DataField       =   "SH_SALARY"
         Height          =   285
         Left            =   1670
         TabIndex        =   7
         Tag             =   "21-Enter salary"
         Top             =   195
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPremium 
         Height          =   285
         Left            =   1665
         TabIndex        =   8
         Tag             =   "21-Enter salary"
         Top             =   750
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTotal 
         Height          =   285
         Left            =   1665
         TabIndex        =   9
         Tag             =   "21-Enter salary"
         Top             =   1185
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblTitle 
         Caption         =   "Premium"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Caption         =   "Total"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Per"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SalCode"
         DataField       =   "SH_SALCD"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8160
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblSalaryGrade 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SH_GRADE"
         DataField       =   "SH_GRADE"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7080
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   5310
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         Caption         =   "Vailtech Group"
         Height          =   255
         Index           =   19
         Left            =   3480
         TabIndex        =   11
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Caption         =   "Vailtech Step"
         Height          =   255
         Index           =   20
         Left            =   3480
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
   End
   Begin INFOHR_Controls.CodeLookup clpGrid 
      DataField       =   "SH_GRID"
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGD"
      TABLTitle       =   "Grid Category"
      MaxLength       =   10
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   8745
      Width           =   10635
      _Version        =   65536
      _ExtentX        =   18759
      _ExtentY        =   741
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
      Begin VB.CommandButton cmdPosition 
         Appearance      =   0  'Flat
         Caption         =   "P&osition"
         Height          =   280
         Left            =   10350
         TabIndex        =   25
         Top             =   330
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.CommandButton cmdPerform 
         Appearance      =   0  'Flat
         Caption         =   "Perfor&mance"
         Height          =   280
         Left            =   10350
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.CommandButton cmdRecal 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Tag             =   "Recalculate Percentage Change"
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdChPos 
         Appearance      =   0  'Flat
         Caption         =   "Edit Position/&Date"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Tag             =   "Edit Position Code and Start Date"
         Top             =   0
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   8250
         Top             =   30
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   2
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
   Begin Threed.SSPanel panEEDesc 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10635
      _Version        =   65536
      _ExtentX        =   18759
      _ExtentY        =   1032
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
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3120
         TabIndex        =   29
         Top             =   180
         Width           =   1740
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1560
         TabIndex        =   28
         Top             =   180
         Width           =   1245
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
         Left            =   270
         TabIndex        =   27
         Top             =   203
         Width           =   1005
      End
   End
   Begin INFOHR_Controls.DateLookup dlpPosStDate 
      DataField       =   "SH_SDATE"
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      Tag             =   "41-Enter Position Start Date"
      Top             =   2430
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmESALARY1.frx":0000
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "frmESALARY1.frx":0014
      TabIndex        =   31
      Top             =   600
      Width           =   9615
   End
   Begin INFOHR_Controls.CodeLookup clpPostCode 
      DataField       =   "SH_JOB"
      Height          =   285
      Left            =   1800
      TabIndex        =   32
      Tag             =   "01-Position code"
      Top             =   2100
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SH_SREAS1"
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   33
      Tag             =   "01-Reason code "
      Top             =   3390
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   390
      Left            =   7080
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
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
      Caption         =   "HREMP"
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
   Begin Threed.SSCheck chkCurrent 
      DataField       =   "SH_CURRENT"
      Height          =   255
      Left            =   6960
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1890
      _Version        =   65536
      _ExtentX        =   3334
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Current Salary Record"
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
      Enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox medPercentChng 
      DataField       =   "SH_SALPC1"
      Height          =   285
      Index           =   1
      Left            =   5370
      TabIndex        =   38
      Tag             =   "10-Percentage change from previous salary"
      Top             =   3390
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPercentChng 
      DataField       =   "SH_SALPC2"
      Height          =   285
      Index           =   2
      Left            =   5370
      TabIndex        =   39
      Tag             =   "10-Percentage change from previous salary"
      Top             =   3705
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPercentChng 
      DataField       =   "SH_SALPC3"
      Height          =   285
      Index           =   3
      Left            =   5370
      TabIndex        =   40
      Tag             =   "10-Percentage change from previous salary"
      Top             =   4020
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmtChng 
      DataField       =   "SH_SALCHG1"
      Height          =   285
      Index           =   1
      Left            =   7290
      TabIndex        =   41
      Tag             =   "20-$ change from previous salary"
      Top             =   3390
      Width           =   1575
      _ExtentX        =   2778
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmtChng 
      DataField       =   "SH_SALCHG2"
      Height          =   285
      Index           =   2
      Left            =   7290
      TabIndex        =   42
      Tag             =   "20-$ change from previous salary"
      Top             =   3705
      Width           =   1575
      _ExtentX        =   2778
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmtChng 
      DataField       =   "SH_SALCHG3"
      Height          =   285
      Index           =   3
      Left            =   7290
      TabIndex        =   43
      Tag             =   "20-$ change from previous salary"
      Top             =   4020
      Width           =   1575
      _ExtentX        =   2778
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
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10560
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "c:\ihr\rgridsal.rpt"
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SH_SREAS2"
      Height          =   285
      Index           =   2
      Left            =   300
      TabIndex        =   44
      Tag             =   "01-Reason code "
      Top             =   3720
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SH_SREAS3"
      Height          =   285
      Index           =   3
      Left            =   300
      TabIndex        =   45
      Tag             =   "01-Reason code "
      Top             =   4050
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   " "
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   46
      Tag             =   "00-Enter Union Code"
      Top             =   9000
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin VB.Label LabelPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   98
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "SH_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   97
      Top             =   8340
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EEId"
      DataField       =   "SH_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   96
      Top             =   8790
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   95
      Top             =   2490
      Width           =   1620
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason For Salary Change"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   94
      Top             =   3150
      Width           =   2280
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage Change"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   5130
      TabIndex        =   93
      Top             =   3150
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Change"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   7050
      TabIndex        =   92
      Top             =   3150
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours per Week:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   3600
      TabIndex        =   91
      Top             =   2490
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Per Pay :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   6960
      TabIndex        =   90
      Top             =   2505
      Width           =   1455
   End
   Begin VB.Label lblPayPeriodSalary 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8400
      TabIndex        =   89
      Top             =   2505
      Width           =   75
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   88
      Top             =   2790
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblBand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Band"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   87
      Top             =   4620
      Width           =   375
   End
   Begin VB.Label lblBANDCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Disp"
      DataField       =   "SH_Band"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10440
      TabIndex        =   86
      Top             =   4620
      Width           =   315
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hourly Rate :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   21
      Left            =   6960
      TabIndex        =   85
      Top             =   2805
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblHoursPay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8160
      TabIndex        =   84
      Top             =   2805
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmESALARY11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fglbEmptyNew
Dim UnionExecNone As Boolean

Dim orgSalary As Double
Dim orgSalary1 As Double
Dim OSalary, OSalCD, OJOB, OEDate, ONDate, OReason
Dim OPremium, OTotal, OvGroup, OVStep 'Vailtech
Dim oGrade
Dim Actn
Dim orgCurrent
Dim SavPAYP, OldPAYP, SavSalcd
Dim orgPosStDate As String
Dim dynaJobHIS As New ADODB.Recordset
Dim fglbJob$, fglbJobID&, fglbReason$
Dim fglbGrid$
Dim fglbPayrollID
Dim fglbSDate, fglbWhrs#, fglbBAND
Dim fglbPhrs, fglbDhrs
Dim OLambtonJob
Dim JobSnaps_PayScale(11) As Double
Dim JobSnaps_Salary_Code$
Dim JobSnaps_Salary_FTEHrs
Dim JobSnap_MidPoint!
Dim fSection As String


Dim fglbPCOld(4) As Double
Dim fglbAmtOld(6) As Currency
Dim fglbSHold@
Dim fglbGridType
Dim fglHredsem As String
Dim fglbNew As Boolean
Dim fglbFrmt As String
Dim flagFrmLoad As Boolean
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbJobList
Dim flgloaded As Boolean
Dim prompt As Boolean
Dim MailBody

Private Function AUDITSALY(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim SQLQ As String, strFields As String

On Error GoTo AUDIT_ERR

AUDITSALY = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    xPT = rsTB("ED_PT")
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
strFields = strFields & "AU_JOB, AU_SEDATE, AU_SREASON, AU_PAYP, AU_OLDPAYP, "
strFields = strFields & "AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_JOB "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False
'~~~~~~~~~CHECK FOR NULL VALUES~~RAUBREY 6/19/97~~~~~~~~~~~
If IsNull(OSalary) Then
    OSalary = 0
End If

If IsNull(ONDate) Then
    ONDate = ""
Else
    If ONDate <> "01/01/01" Then    'THIS IS TO ENSURE THAT ONDATE
        ONDate = Trim(str$(ONDate)) '    HAS NOT ALREADY BEEN SET
    End If                          '    TO A STRING IN Function
End If                              '    CurSHDate
If glbVadim And Not IsDate(ONDate) Then
    ONDate = "01/01/01"
End If

If IsNull(OEDate) Then
    OEDate = ""
Else
    If OEDate <> "01/01/01" Then 'THIS IS TO ENSURE THAT OEDATE HAS NOT ALREADY BEEN SET TO A STRING IN Function CurSHDate
        OEDate = Trim(str$(OEDate))
    End If
End If

'do not know what should we do if there is salary changes

Dim xBatchID, UpdateAudit
Dim HRChanges As New Collection
Dim UptSalaryDate As Date
If fglbNew Then
    UptSalaryDate = dlpDate(0)
Else
    UptSalaryDate = Date
End If
UpdateAudit = False
Dim HRSalary As New Collection
'Town of Aurora or City of Timmins or City of Niagara Falls
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then 'Or glbCompSerial = "S/N - 2363W" Then
    'Or glbCompSerial = "S/N - 2276W" Then
    
    Dim LowSalary As New FieldInfo
    Dim LowSalCD As New FieldInfo
    Dim LowGrade As New FieldInfo
    Dim LowEDate As New FieldInfo
    Dim LowNDate As New FieldInfo
    Dim LowPAYP As New FieldInfo
    Dim LowReason As New FieldInfo
    Dim LowJob As New FieldInfo
    Dim LowDHRS As New FieldInfo
    Dim rsOldSalary As New ADODB.Recordset
    SQLQ = "Select * from HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & " AND SH_CURRENT <>0"
    If Not fglbNew Then
        SQLQ = SQLQ & " AND SH_ID<>" & Data1.Recordset("SH_ID")
    End If
    SQLQ = SQLQ & " ORDER BY SH_SALARY"
    rsOldSalary.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If chkCurrent = 0 Then
        LowSalary.fdValue = 0: LowSalary.fdName = medsalary.DataField
    Else
        LowSalary.fdValue = Val(medsalary): LowSalary.fdName = medsalary.DataField
    End If
    LowSalCD.fdValue = lblSalCode: LowSalCD.fdName = lblSalCode.DataField
    LowGrade.fdValue = lblSalaryGrade: LowGrade.fdName = lblSalaryGrade.DataField
    LowEDate.fdValue = dlpDate(0): LowEDate.fdName = dlpDate(0).DataField
    LowNDate.fdValue = dlpDate(1): LowNDate.fdName = dlpDate(1).DataField
    LowPAYP.fdValue = clpCode(4): LowPAYP.fdName = clpCode(4).DataField
    LowReason.fdValue = clpCode(1): LowReason.fdName = clpCode(1).DataField
    LowJob.fdValue = clpPostCode: LowJob.fdName = "JH_JOB"
    LowDHRS.fdValue = 0: LowDHRS.fdName = "JH_DHRS"

    If Not rsOldSalary.EOF Then
        If rsOldSalary("SH_SALARY") < LowSalary.fdValue Or chkCurrent = 0 Then
            LowSalary.fdValue = rsOldSalary("SH_SALARY")
            LowSalCD.fdValue = rsOldSalary("SH_SALCD")
            LowGrade.fdValue = rsOldSalary("SH_GRADE")
            LowEDate.fdValue = rsOldSalary("SH_EDATE")
            LowNDate.fdValue = rsOldSalary("SH_NEXTDAT")
            LowPAYP.fdValue = rsOldSalary("SH_PAYP")
            LowReason.fdValue = rsOldSalary("SH_SREAS1")
            LowJob.fdValue = rsOldSalary("SH_JOB")
        End If
    End If
    If isChanged_Salary(HRSalary, OSalary, LowSalary, True) Then UpdateAudit = True
    If isChanged_Salary(HRSalary, OSalCD, LowSalCD) Then UpdateAudit = True
    If glbVadim And UpdateAudit Then
        'City of Niagara Falls has special logic to calculate the Hourly Rate so send
        'Hours per Day instead of Hours per Week
        If glbCompSerial = "S/N - 2276W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbDhrs, glbLEE_ID, txtPayrollID.Text)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbWhrs, glbLEE_ID, txtPayrollID.Text)
        End If
        If isChanged_Field(HRChanges, oGrade, LowGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
    End If
    
    
    If isChanged_Field(HRChanges, OEDate, LowEDate) Then UpdateAudit = True
    If isChanged_Field(HRChanges, ONDate, LowNDate) Then UpdateAudit = True
    If isChanged_Field(HRChanges, SavPAYP, LowPAYP) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OReason, LowReason) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OJOB, LowJob) Then UpdateAudit = True
    If OJOB <> LowJob.fdValue Then
        rsTC.Open "SELECT JH_DHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_JOB='" & LowJob.fdValue & "' AND JH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
        If Not rsTC.EOF Then
            LowDHRS.fdValue = rsTC("JH_DHRS")
        End If
        rsTC.Close
        If isChanged_Field(HRChanges, 0, LowDHRS) Then UpdateAudit = True
    End If
    Call Passing_Changes(HRChanges, Salary, "M", Date, glbLEE_ID, txtPayrollID.Text)

Else
    
    If isChanged_Salary(HRSalary, OSalary, medsalary, True) Then UpdateAudit = True
    If isChanged_Salary(HRSalary, OSalCD, lblSalCode) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbWhrs, glbLEE_ID, txtPayrollID.Text)
        'City of Kawartha Lakes - Pass Salary Grade to Probation Levels
        If glbCompSerial = "S/N - 2363W" Then
            If fglbNew Then
                If isChanged_Field(HRChanges, "", lblSalaryGrade, True) Then UpdateAudit = True
            Else
                If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then UpdateAudit = True
            End If
        Else
            If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
        End If
    End If
    
    
    If isChanged_Field(HRChanges, OEDate, dlpDate(0)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, ONDate, dlpDate(1)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, SavPAYP, clpCode(4)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OReason, clpCode(1)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OJOB, clpPostCode) Then UpdateAudit = True
    If glbCompSerial = "S/N - 2373W" Then 'DMuskoka , ,  'Vailtech
        If isChanged_Field(HRChanges, OPremium, medPremium) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OTotal, medTotal) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OvGroup, txtVGroup) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OVStep, txtVStep) Then UpdateAudit = True
    End If
    Call Passing_Changes(HRChanges, Salary, "M", Date, glbLEE_ID, txtPayrollID.Text)
End If
If UpdateAudit Then GoTo MODUPD Else GoTo MODNOUPD


MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_GRID") = clpGrid.Text


If Trim(str$(OSalary)) <> medsalary Or SavSalcd <> lblSalCode Then 'Trim(Str$ added by RAUBREY 6/3/97
    rsTA("AU_SALARY") = medsalary
    rsTA("AU_OLDSAL") = OSalary
    rsTA("AU_WHRS") = lblWhrs
    rsTA("AU_SALCD") = lblSalCode 'laura febr 2, 1998
End If

'Added by Bryan 27/09/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'muskoka
    If (OPremium) <> medPremium Or (OTotal) <> medTotal Or OvGroup <> txtVGroup Or OVStep <> txtVStep Then
        rsTA("AU_TOTAL") = medTotal
        rsTA("AU_VPREMIUM") = medPremium
        rsTA("AU_VGROUP") = txtVGroup
        rsTA("AU_VSTEP") = txtVStep
    End If
End If
    

If glbInsync Then
    rsTA("AU_JOB") = clpPostCode.Text
    rsTA("AU_SEDATE") = dlpDate(0).Text
    rsTA("AU_SREASON") = clpCode(1).Text
Else
    If OJOB <> clpPostCode.Text Then rsTA("AU_JOB") = clpPostCode.Text
    If OEDate <> dlpDate(0).Text Then rsTA("AU_SEDATE") = dlpDate(0).Text
    If OReason <> clpCode(1).Text Then rsTA("AU_SREASON") = clpCode(1).Text
End If

If SavPAYP <> clpCode(4).Text Then
    If Len(clpCode(4).Text) > 0 Then
        rsTA("AU_PAYP") = clpCode(4).Text
    Else
        rsTA("AU_PAYP") = "-"
    End If
    If Not IsNull(SavPAYP) Then
        If SavPAYP <> "" Then rsTA("AU_OLDPAYP") = SavPAYP
    End If
Else
    If Val(clpCode(4).Text) = 0 Then
        rsTA("AU_PAYP") = Null
    Else
        rsTA("AU_PAYP") = Val(clpCode(4).Text)
    End If
   If SavPAYP <> "" Then rsTA("AU_OLDPAYP") = Val(SavPAYP)
End If


If IsDate(dlpDate(1).Text) Then                   '13Aug99 js
    If ONDate <> dlpDate(1).Text Then             '
        rsTA("AU_SNDATE") = dlpDate(1).Text      '
    End If                                  '
Else                                        '
    rsTA("AU_SNDATE") = Null                '
End If                                      '

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID


If glbCompSerial = "S/N - 2290W" Then
    rsTA("AU_LDATE") = Date
Else
    If Actn = "A" Then
        If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
            rsTA("AU_LDATE") = Format(DateAdd("d", 14, dlpDate(0)), "SHORT DATE")
        Else
            rsTA("AU_LDATE") = dlpDate(0).Text
        End If
    Else
        If dlpDate(0) > Date Then
            rsTA("AU_LDATE") = dlpDate(0)
        Else
            rsTA("AU_LDATE") = Date
        End If
    End If
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
If glbMulti Then
    rsTA("AU_PAYROLL_ID") = txtPayrollID
Else
    Dim rsEmp As New ADODB.Recordset
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    If OSalary <> medsalary Then
        rsTA("AU_JOB") = clpPostCode.Text '# 7644
    End If
End If
rsTA.Update
' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
Call Pause(0.5)


'~~~~~~~~~~~~~~~~~~~~~~~~

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
Private Sub TermRehireAudit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTilPayID
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If rsTC.EOF Then Exit Sub
    'If IsNull(rsTC("ED_PAYROLL_ID")) Then Exit Sub
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
    rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
    rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
    rsTA("AU_PAYP") = OldPAYP
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_VERSION") = "ADPTRA" 'Ticket# 7768
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    rsTA.Update
    
    'New Hire Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1":: rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_DIV") = rsTC("ED_DIV")
    rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
    rsTA("AU_TITLE") = rsTC("ED_TITLE")
    rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
    rsTA("AU_FNAME") = rsTC("ED_FNAME")
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
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



Private Sub cboVGRoup_Click()
txtVGroup = cboVGRoup.Text
End Sub

Private Sub cboVStep_Click()
txtVStep = cboVStep.Text
End Sub

Private Sub chkCurrent_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function Chkpos()
Dim SQLQ As String, Msg$
Dim xPosFind

On Error GoTo ChkPos_Err

Chkpos = False

If Len(dlpPosStDate.Text) < 1 Then
    ' If pos. start date is missing in multi, it means they didn't enter a valid position
    If glbMulti Then
        MsgBox "Position does not exist in Position History file.  Please correct this before continuing.", vbOKOnly + vbExclamation, "Position Not Found"
         clpPostCode.SetFocus
    Else
        Msg$ = "Position Start Date is required"
        dlpPosStDate.SetFocus
        MsgBox Msg$
    End If
    Exit Function
Else
    If Not IsDate(dlpPosStDate.Text) Then
        Msg$ = "Not a Valid Position Start Date"
        dlpPosStDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If

End If

If Len(dlpDate(0).Text) < 1 Then
    Msg$ = "Effective Date is required"
    dlpDate(0).SetFocus
    MsgBox Msg$
    Exit Function
End If

If Len(clpPostCode.Text) > 0 Then
    If clpPostCode.Caption = "Unassigned" Then
        MsgBox "Position Code is invalid"
         clpPostCode.SetFocus
        Exit Function
    End If
Else
    If clpPostCode.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
         clpPostCode.SetFocus
        Exit Function
    End If
End If
xPosFind = False
If Not Set_Position(clpPostCode.Text, False) Then
    Msg$ = "No position <" & clpPostCode.Text & "> found "
    Msg$ = Msg$ & Chr(10) & "Please review positions from Position History!"
    MsgBox Msg$
    Exit Function
End If
If dlpPosStDate.Text <> fglbSDate Then
    MsgBox "Start Date in the Salary History is different than the Position History!"
End If
If glbMultiGrid Then
    If Len(clpGrid.Text) <= 0 Then
        MsgBox lStr("Grid Category is required")
        clpGrid.SetFocus
        Exit Function
    Else
        If clpGrid.Caption = "Unassigned" Then
            MsgBox lStr("Grid Category is required")
            clpGrid.SetFocus
            Exit Function
        End If
    End If
End If
If glbMulti And glbVadim Then
    Dim rsChkJob As New ADODB.Recordset
    If chkCurrent Then
        rsChkJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID & " AND JH_PAYROLL_ID='" & txtPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
        If rsChkJob.EOF Then
            Msg$ = "No Payroll ID found in the Current Positions"
            Msg$ = Msg$ & Chr(10) & "Please review positions from Position History!"
            MsgBox Msg$
            txtPayrollID.SetFocus
            Exit Function
        End If
        rsChkJob.Close
    End If
End If
Chkpos = True

Exit Function

ChkPos_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdChPos", "HR_JOB_HISTORY", "Change Position")
Resume Next

End Function

Private Function chkSalHist()
Dim X%
Dim SQLQ As String, Msg$, dd&
Dim DgDef As Variant, Title$, Response%, DCurSHDate  As Variant
Dim rsEmp As New ADODB.Recordset
Dim dtEmpDOH As Date
chkSalHist = False

On Error GoTo chkSalH_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Reason Code is required"
    clpCode(1).SetFocus
    Exit Function
Else
    For X% = 1 To 4
        If X% < 4 Then
            If Len(clpCode(X%).Text) = 0 Then
                medPercentChng(X%) = 0
                medAmtChng(X%) = 0
            End If
        End If
        If clpCode(X%).Caption = "Unassigned" Then
            If X% < 4 Then
                MsgBox "Reason Code must be valid"
            Else
                MsgBox "Pay Period Code must be valid"
            End If
            clpCode(X%).SetFocus
            Exit Function
        End If
    Next X%
End If
If glbVadim Then
    If glbMulti Then 'Ticket# 7751
        If Len(txtPayrollID.Text) = 0 Then
            MsgBox "Payroll ID is required"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
End If

If glbPayWeb Or glbVadim Or glbLambton Or glbInsync Or glbCompSerial = "S/N - 2348W" _
Or glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2370W" _
Or (glbWFC And fSection = "GREN") Then
    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Pay Period Code is required"
        clpCode(4).SetFocus
        Exit Function
    End If
    
End If
' -----

If Len(medsalary) < 1 Then
    If fraSalary.Enabled = True Then medsalary.SetFocus
    MsgBox "Salary is required"
    If medsalary.Enabled Then medsalary.SetFocus
    Exit Function
End If
If medsalary <= 0 Then
    If fraSalary.Enabled = True Then medsalary.SetFocus
    MsgBox "Salary is required"
    If medsalary.Enabled Then medsalary.SetFocus
    Exit Function
End If
' -----
'Hemu - 06/18/2003 Begin - Incase the 'Per' has no value
    If comPayPer.Text = "" Then
        MsgBox "Per cannot be blank"
        comPayPer.SetFocus
        Exit Function
    End If

'Hemu - 06/18/2003 End

If glbWFC Then 'Frank 09/24/04 Ticket# 6962
    If clpCode(0).Visible And Len(clpCode(0).Text) < 1 Then
        Msg$ = "Plant is required"
        clpCode(0).SetFocus
        MsgBox Msg$
        Exit Function
    End If
    If txtFiscalYear.Visible And Len(txtFiscalYear) < 1 Then
        Msg$ = "Fiscal Year is required"
        txtFiscalYear.SetFocus
        MsgBox Msg$
        Exit Function
    End If
    If cmbMarketLine.Visible And Len(cmbMarketLine.Text) < 1 Then
        Msg$ = "Market Line is required"
        cmbMarketLine.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If Len(dlpDate(0).Text) < 1 Then
    Msg$ = "Effective Date is required"
    dlpDate(0).SetFocus
    MsgBox Msg$
    Exit Function
Else
    If Not IsDate(dlpDate(0).Text) Then
        Msg$ = "Not a Valid Effective Date"
        dlpDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        DCurSHDate = CurSHDate()
        If DCurSHDate > 0 Then    ' 0 if no current record out there
           DCurSHDate = CVDate(DCurSHDate)
           If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <> 0 Then
                Call ChangeEDateAudit(DCurSHDate)
                
           End If
        End If
        If glbSetSal Then
            If DCurSHDate > 0 Then    ' 0 if no current record out there
                DCurSHDate = CVDate(DCurSHDate)
                If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <= 0 Then
                    Msg$ = "Warning...you cannot add or edit a record with a date"
                    Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current salary, "
                    Msg$ = Msg$ & Chr(10) & "go to Salary screen under Employee Menu."
                    MsgBox Msg$
                    dlpDate(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If
'Hemu 05/13/2003 Begin - Effective Date and Original Hire Date
If Len(dlpDate(0).Text) > 0 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Effective Date is not a valid date"
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Not glbLambton Then
        rsEmp.Open "SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            If rsEmp("ED_DOH") <> "" Then
            
            dtEmpDOH = rsEmp("ED_DOH")
            If DaysBetween(rsEmp("ED_DOH"), dlpDate(0).Text) < 0 Then
                MsgBox "Effective Date can not be prior to Original Hire date"
                dlpDate(0).SetFocus
                rsEmp.Close
                Exit Function
            End If
            End If
        End If
        rsEmp.Close
    End If
End If
'Hemu 05/13/2003 End

DCurSHDate = CurSHDate()
If Not fglbNew And glbMediPay Then
    Dim OtherChange
    If SavPAYP <> clpCode(4) Then
        OtherChange = False

        If CDbl(OSalary) <> CDbl(medsalary) Then OtherChange = True
        If OSalCD <> lblSalCode Then OtherChange = True
        If OEDate <> dlpDate(0) Then OtherChange = True
        If ONDate <> dlpDate(1) Then OtherChange = True
        If OReason <> clpCode(1) Then OtherChange = True
        If OJOB <> clpPostCode Then OtherChange = True
        If OtherChange Then
            Msg$ = "Warning, you can not change Salary information with the Client # transfer."
            Msg$ = Msg$ & Chr(10) & "Please cancel the changes."
            DgDef = MB_OK + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg$) ', DgDef, "Warning!")
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
End If
If glbAddHisWarning And Actn = "A" And (Not glbSetSal) Then
    If DCurSHDate > 0 Then    ' 0 if no current record out there
        DCurSHDate = CVDate(DCurSHDate)
        If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) > 0 Then
            Msg$ = "Warning, you can not add a record with a date"
            Msg$ = Msg$ & Chr(10) & "earlier than your most current record."
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg$) ', DgDef, "Warning!")
            dlpDate(0).SetFocus
            Exit Function
        End If
    End If
End If

If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Next Review Date is invalid"
        dlpDate(1).SetFocus
        Exit Function
    End If
        'Hemu - 05/13/2003 Begin
    If DaysBetween(dtEmpDOH, dlpDate(1).Text) < 0 Then
        MsgBox "Next Review date can not be prior to Original Hire date"
        dlpDate(1).SetFocus
        Exit Function
    End If
    'Hemu - 05/13/2003 End
    dd& = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    If dd& < 0 Then
        Msg$ = "Next Review precedes Effective date of salary "
        dlpDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    End If
Else
    If glbLinamar And (chkCurrent Or Actn = "A") Then
        MsgBox "Next Review Date is required"
        dlpDate(1).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(2).Text) > 0 Then
    If Not IsDate(dlpDate(2).Text) Then
        MsgBox "Transaction Date is invalid"
        dlpDate(2).SetFocus
        Exit Function
    End If
End If


' dkostka - 03/20/2002 - Added check for user compa box
' if it's woodbridge and compa is set by user, they have to enter a value.
'Ticket# 6962 don't need this any more 09/24/04 Frank
'If glbWFC And optUserSys(1).Value = True And (mskCampa.Text = "" Or Not IsNumeric(mskCampa.Text)) Then
'    MsgBox "If Compa Ratio is set to user, a value must be entered.", vbExclamation + vbOKOnly, "Value Required"
'    mskCampa.SetFocus
'    Exit Function
'End If



'Frank 08/27/03 - Pay Period is mandatory for Soroc
If glbSoroc Or glbSyndesis Or glbCompSerial = "S/N - 2229W" Then 'Soroc, Syndesis,Inscape
    If Len(clpCode(4).Text) < 1 Then
        Msg$ = "Pay Period is required"
        clpCode(4).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If (glbCompSerial = "S/N - 2242W") Then  'C.C.A.C. London & Middlesex - Ticket #6718
    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Client # is required"
        clpCode(4).SetFocus
        Exit Function
    End If
    
    If Not clpCode(4).ListChecker Then
        MsgBox "Client # must be valid"
        clpCode(4).SetFocus
    End If
End If

If DCurSHDate = 0 Then DCurSHDate = dlpDate(0).Text   'New Record
If IsDate(DCurSHDate) Then
    If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <= 0 Then
        If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then  'Town of Aurora and Timmmins
            If Not AUDITSALY(Actn) Then MsgBox "ERROR - AUDIT FILE"
        ElseIf Not glbMulti Or chkCurrent = True Then
            If Not AUDITSALY(Actn) Then MsgBox "ERROR - AUDIT FILE"
        End If
    End If
End If

chkSalHist = True

Exit Function

chkSalH_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSal", "HR_SALARY_HISTORY", "edit/Add")
Resume Next

End Function

Private Sub clpCode_LostFocus(Index As Integer)
If Index = 0 Then
    Call Set_SalState
    Call Set_MarketLine_List
End If
If Index = 5 Then
    txtComment = clpCode(5)
End If
End Sub

Private Sub clpGrid_LostFocus()
If Len(clpPostCode) = 0 Then Exit Sub
Call getJOB(clpPostCode, clpGrid)
If Set_Position(clpPostCode, False) Then
End If
If glbMulti Then Call Get_OrgSalary
End Sub

Private Sub clpPostCode_LostFocus()
If Len(clpPostCode) = 0 Then Exit Sub
If Set_Position(clpPostCode, False) Then
    lblBANDCode = fglbBAND
    dlpPosStDate = fglbSDate
    clpGrid = fglbGrid
    txtWHRS = fglbWhrs
    txtPayrollID = fglbPayrollID
Else
    lblBANDCode = ""
    dlpPosStDate = ""
    clpGrid = ""
    txtWHRS = ""
    txtPayrollID = ""
End If
Call getJOB(clpPostCode, clpGrid)

If glbMulti Then Call Get_OrgSalary
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err

dlpDate(0).DataChanged = False
dlpDate(1).DataChanged = False
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
'rsDATA.CancelUpdate

fglbNew = False
Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_SALARY_HISTORY", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdChPos_Click()
clpPostCode.Enabled = True
dlpPosStDate.Enabled = True
clpGrid.Enabled = True
clpPostCode.SetFocus
If chkCurrent.Value = 0 Then txtWHRS.Enabled = True Else txtWHRS.Enabled = False
End Sub

Private Sub cmdChPos_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESALARY11" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, xID
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant
Dim DCurSHDate
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then  'Town of Aurora and timmis
    If chkCurrent <> 0 Then
        MsgBox "Please uncheck the Current Salary flag before deleting the record"
        Exit Sub
    End If
End If
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

DtTm = Now
DCurSHDate = CurSHDate()
fglHredsem = dlpDate(1).Text  '11/2/97 by Laura
If fglHredsem <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

ResetFlagAudit

xID = Data1.Recordset("SH_ID")
gdbAdoIhr001.BeginTrans
rsDATA.Delete 'gdbAdoIhr001.Execute "DELETE FROM HR_SALARY_HISTORY WHERE SH_ID=" & xID
gdbAdoIhr001.CommitTrans
If Not glbOracle And Not glbSQL Then Pause (0.5)
Data1.Refresh

If glbGP Then Call Salary_Integration(glbLEE_ID, , True, fglbNew, xID) 'George Mar 7,2006 #9965

prompt = False
Call cmdRecal_Click
prompt = True
Data1.Refresh

If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    Call Set_Current_Flag
End If

Call Display_Value

If OSalary <> medsalary And (chkCurrent Or Data1.Recordset.EOF) Then
    Call updBenefitForSalDEPN(glbLEE_ID) 'Jaddy 9/9/99
    If glbCompSerial = "S/N - 2291W" Then Call updCompPlan(glbLEE_ID, Val(medsalary) - Val(OSalary), DCurSHDate)
End If
Call Employee_Master_Integration(glbLEE_ID)
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_SALARY_HISTORY", "Delete")
Call RollBack '28July99 js
End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim SQLQ As String, X%
Dim Response%, Msg$, Title$, DgDef As Double

On Error GoTo Mod_Err

Call SET_UP_MODE

Actn = "M"
fglHredsem = dlpDate(1).Text
If Not Data1.Recordset.EOF Then
    If Not IsNull(Data1.Recordset("SH_JOB")) Then
        fglbJob$ = Data1.Recordset("SH_JOB")
    End If
End If
orgPosStDate = dlpPosStDate.Text

'orgSalary = Val(medSalary)
orgSalary1 = Val(medsalary)

'Hemu - essex
fglbAmtOld(1) = CCur(Val(medAmtChng(1)))
fglbAmtOld(2) = CCur(Val(medAmtChng(2)))
fglbAmtOld(3) = CCur(Val(medAmtChng(3)))
'Hemu - essex

orgCurrent = chkCurrent
SavPAYP = clpCode(4).Text


SavSalcd = lblSalCode


''If glbWFC And UnionExecNone Then
''    lblBANDCode = fglbBAND
''    optUserSys(0).Value = False: optUserSys(1).Value = True
''    optUserSys(0).Enabled = False: optUserSys(1).Enabled = True
''    mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible
''    If Val(lblsalstate(1)) > 0 And Val(mskCampa) = 0 Then
''      If Val(lblCompaNum) > 0 And Val(lblCompaNum) < 999.99 Then
''        mskCampa = (Val(medSalary) / Val(lblCompaNum)) * 100
''      Else
''        mskCampa = Val(lblsalstate(1))
''      End If
''      mskCampa = Round2DEC(mskCampa)
''    End If
''End If

'clpCode(1).SetFocus
'If glbSetSal Or glbMulti Then clpPostCode.SetFocus

'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_SALARY_HISTORY", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$
Dim X%
Dim orgMarketLine, orgSalCD
Dim xPayPeriod
On Error GoTo AddN_Err
fglbNew = True

'Hemu - essex
fglbAmtOld(1) = 0
fglbAmtOld(2) = 0
fglbAmtOld(3) = 0
'Hemu - essex

Call CR_JobHis_Snap
If Not Set_Position("", True) Then
    Msg$ = "No current position found "
    Msg$ = Msg$ & Chr(10) & "Please review position prior to updating salary."
    MsgBox Msg$
    Exit Sub
End If
If Not getJOB(fglbJob$, fglbGrid) Then   '- populates job items/grades
    If glbMultiGrid Then
        Msg$ = "Can not find Salary Details for current position and grid category."
        Msg$ = Msg$ & Chr(10) & "Please review position Master list and the Salary Details."
    Else
        Msg$ = "Can not find description for current position."
        Msg$ = Msg$ & Chr(10) & "Please review position Master list."
    End If
    MsgBox Msg$
    Exit Sub
End If
If glbMulti And Not Data1.Recordset.EOF Then
    MsgBox "If necessary, edit the previous salary record to remove the current flag."
End If
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    Data1.Recordset.MoveFirst
    orgMarketLine = txtMarketLine
    orgSalary = Val(medsalary)
    orgSalary1 = Val(medsalary)
    orgSalCD = lblSalCode

    If glbMulti Then Call Get_OrgSalary
Else
    orgMarketLine = ""
    orgSalary = 0
    orgSalary1 = 0
    orgSalCD = JobSnaps_Salary_Code$
End If
DoEvents
xPayPeriod = clpCode(4)
fglbEmptyNew = (Data1.Recordset.BOF And Data1.Recordset.EOF)
Call Set_Control("B", Me)

'rsDATA.AddNew

If fglbReason$ = "NEWH" And fglbEmptyNew Then clpCode(1).Text = "NEWH"

Actn = "A"


lblCNum.Caption = "001"
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblWhrs = fglbWhrs#
txtWHRS = fglbWhrs#

clpPostCode.Text = fglbJob$
dlpPosStDate.Text = CVDate(fglbSDate)
clpGrid.Text = fglbGrid
txtPayrollID = fglbPayrollID
lblBANDCode = fglbBAND
Call setGridList(fglbJob$)
'If glbLinNewPosSal And glbLinamar Then 'Jaddy changed by linda asking, 8/20/01
If glbLinamar Then
    Call Set_NextReview
    If glbLinNewPosSal Then
        clpCode(1).Text = fglbReason$  'glbLinReasonCode
    End If
End If

If glbLambton Then
    
    If Len(xPayPeriod) > 0 Then
        clpCode(4) = xPayPeriod
    Else
        clpCode(4) = "26"
    End If
End If
If glbMediPay Then
    If Len(xPayPeriod) > 0 Then
        clpCode(4) = xPayPeriod
    End If
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    Call Set_CommentFromUnion
End If

If glbCompSerial = "S/N - 2229W" Then 'Inscap Solution - Ticket # 8932
    If Len(xPayPeriod) > 0 Then
        clpCode(4) = xPayPeriod
    End If
End If

lblSalaryGrade = "00"



lblSalCode = orgSalCD
chkCurrent = glbMulti
medsalary = 0
SavPAYP = ""
SavSalcd = ""


fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
clpCode(1).Enabled = True
clpCode(1).SetFocus

If clpCode(1).Text = "NEWH" Then

    fraSalary.Enabled = True
    For X% = 1 To 3
        medPercentChng(X%) = 0
        medPercentChng(X%).Enabled = False
        medAmtChng(X%) = 0
        medAmtChng(X%).Enabled = False
        If X% > 1 Then
            clpCode(X%).Enabled = False
        End If
    Next X%
Else
    medPercentChng(1).Enabled = True
    medAmtChng(1).Enabled = True
End If
comSalScale.ListIndex = 0
If glbMediPay Then
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        clpCode(4).Enabled = False
    End If
End If
If glbWFC Then
    For X% = 0 To cmbMarketLine.ListCount
        If cmbMarketLine.List(X%) = orgMarketLine Then txtMarketLine = orgMarketLine
    Next
    'Ticket# 6962 Begin
    If clpCode(0).Visible Then
        clpCode(0) = glbEmpPlant
    End If
    If dlpDate(2).Visible Then
        dlpDate(2) = Format(Now, "SHORT DATE")
    End If
    'Ticket# 6962 Begin
End If
'If glbSetSal Or glbMulti Then clpPostCode.SetFocus

DoWFCGrids (True)

''added by Bryan 24/Oct/05 Ticket#9607
'If glbCompSerial = "S/N - 2378W" Then
'    txtPayrollID = glbLEE_ID
'End If

Exit Sub

AddN_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

End Sub
Private Sub Set_CommentFromUnion()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("JH_ORG")) Then
            clpCode(5).Text = rsTemp("JH_ORG")
        End If
    End If
    rsTemp.Close
    
End Sub
'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Function Set_SalaryGrade(xSalary As Double)
Dim SQLQ As String, X%
Dim xsSalary As Double
Dim strSalcode As String
If glbLambton Then 'Ticket# 6693
    If glbSetSal Then
        Exit Function
    End If
End If
If Len(fglbJob$) > 0 Then
    lblSalaryGrade = "00"
    xSalary = Round2DEC(xSalary)
    For X% = 1 To 11
        If JobSnaps_Salary_Code$ = "H" Then
            If lblSalCode = "H" Then
                xsSalary = xSalary
            ElseIf lblSalCode = "M" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    xsSalary = ((xSalary * 12) / Val(lblWhrs)) / 52
                End If
            ElseIf lblSalCode = "A" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        xsSalary = (xSalary)
                    Else
                    xsSalary = (xSalary / Val(lblWhrs)) / 52
                    End If
                End If
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If Val(lblWhrs) = 0 Then
                        xsSalary = 0
                    Else
                        If GetLeapYear(Year(Date)) Then
                            xsSalary = ((xSalary * 366) / Val(lblWhrs)) / 52
                        Else
                            xsSalary = ((xSalary * 365) / Val(lblWhrs)) / 52
                        End If
                    End If
                End If
        ElseIf JobSnaps_Salary_Code$ = "A" Then
            If lblSalCode = "H" Then
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xsSalary = (xSalary)
                Else
                xsSalary = (xSalary * Val(lblWhrs)) * 52
                End If
            ElseIf lblSalCode = "M" Then
                xsSalary = xSalary * 12
            ElseIf lblSalCode = "A" Then
                xsSalary = xSalary
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xsSalary = (xSalary * 366)
                Else
                    xsSalary = (xSalary * 365)
                End If
            End If
        End If
        xsSalary = Round2DEC(xsSalary)
        If JobSnaps_PayScale(X%) <> 0 And xsSalary >= JobSnaps_PayScale(X%) Then
            lblSalaryGrade = Format(X%, "00")
        End If
    Next X%
End If
End Function

Sub cmdOK_Click()
Dim rsSAL As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim X, xID, xUpdCurrent
Dim vList As String
Dim SHMark

On Error GoTo Add_Err

If glbWFC And UnionExecNone Then
    lblBANDCode = fglbBAND
    'optUserSys(0).Value = False: optUserSys(1).Value = True
    'optUserSys(0).Enabled = False: optUserSys(1).Enabled = True
    'mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible
    'If Val(lblsalstate(1)) > 0 And Val(mskCampa) = 0 Then
    '  If Val(lblCompaNum) > 0 And Val(lblCompaNum) < 999.99 Then
    '    mskCampa = (Val(medSalary) / Val(lblCompaNum)) * 100
    '  Else
    '    mskCampa = Val(lblsalstate(1))
    '  End If
    '  mskCampa = Round2DEC(mskCampa)
    'End If
End If

'Hemu - it was not saving the new the Group and Step if new items were added to the list
'commented it here and added these line below - after it assigns the value to txt fields
'vList = VGroupList
'vList = VStepList

If Not chkSalHist() Then Exit Sub

If clpPostCode.Enabled = True Then      'Laura nov 21, 1997
    If Not Chkpos() Then Exit Sub
End If

If gsEMAIL_ONSALARY Then
    MailBody = ""
    If NewHireForms.count = 0 Then 'Non new hire
        If OSalary <> medsalary And (fglbNew Or chkCurrent) Then 'Only Salary Change
            MailBody = "The Salary has been changed." & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            MailBody = MailBody & "New Salary: " & (Format(medsalary, "$#.00")) & "/" & comPayPer.Text & vbCrLf
            MailBody = MailBody & "Reason: " & GetTablDesc("SDRC", clpCode(1)) & vbCrLf
            MailBody = MailBody & "Effective Date: " & dlpDate(0) & vbCrLf
            'Screen.MousePointer = DEFAULT
            'Call imgEmail_Click
        End If
    End If
End If

'If Not chkSalHist() Then Exit Sub
Screen.MousePointer = HOURGLASS

If glbCompSerial = "S/N - 2351W" Then    'Burlington Tech
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
            frmMsgTerm.txtEmpNum.Enabled = False
            frmMsgTerm.Show 1
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

If glbCompSerial = "S/N - 2242W" Then    'London CCAC
    glbChgTermDate = ""
    glbChgPT = ""
    glbChgUseProfile = ""
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
        Select Case SavPAYP
        Case "132"
            frmMsgConfirm.clpCode(0).Text = "CAS"
            frmMsgConfirm.clpCode(1).Text = "NO"
        Case "133"
            frmMsgConfirm.clpCode(0).Text = "FT"
            frmMsgConfirm.clpCode(1).Text = "YES"
        End Select
        frmMsgConfirm.Show 1
        If glbChgPT = "" Or glbChgUseProfile = "" Then
            Call cmdCancel_Click
            Exit Sub
        End If
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

'If this function is processing, it's disabled. ticket #10398
If glbDisabled Then GoTo End_Line
glbDisabled = True

rsDATA.Requery
If fglbNew Then rsDATA.AddNew

Call UpdUStats(Me) ' update user's stats (who did it and when)

If (glbCompSerial = "S/N - 2290W") Or (glbCompSerial = "S/N - 2171W") Then
    Updstats(0).Text = Format(Now, "SHORT DATE")
Else
    Updstats(0).Text = Format(dlpDate(0).Text, "SHORT DATE")
End If

If Not glbWFC Then
    dlpDate(2).Text = Format(Now, "SHORT DATE")
End If
'added by Bryan 22/Sep/05 Ticket#9343
If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
    medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
    medTotal.Text = medsalary.Text
End If
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    Call Set_SalaryGrade(Val(medTotal))
    txtVGroup = cboVGRoup
    txtVStep = cboVStep
Else
    Call Set_SalaryGrade(Val(medsalary))
    
    'City of Timmins
    If glbCompSerial = "S/N - 2375W" Then
        'New Hires don't have a value in comsalscale. Ticket #10436
        Dim strScale As String
        If comSalScale.Text = "" Then
            strScale = 0
        Else
            strScale = comSalScale.Text
        End If
        If JobSnaps_PayScale(CInt(strScale)) <> Val(medsalary) Then
            MsgBox "Salary does not match the grid Step.", vbExclamation, "INFO:HR"
        End If
    End If
    
End If

vList = VGroupList
vList = VStepList

Call Set_COMPA
Call Set_WFC_COMPA

If Actn = "A" Or orgCurrent <> chkCurrent Then
    xUpdCurrent = True
End If

If glbCompSerial = "S/N - 2214W" Then
    Dim xToDate
    If IsDate(dlpDate(1).Text) Then
        xToDate = dlpDate(1)
    Else
        xToDate = DateAdd("D", -1, DateAdd("YYYY", 1, CVDate(dlpDate(0).Text)))
    End If
    If Actn = "A" Then
        Call ChangeOtherEarnAmount(lblEEID, medsalary, "A", dlpDate(0).Text, xToDate)
    End If
    If Actn = "M" And chkCurrent Then
        Call ChangeOtherEarnAmount(lblEEID, medsalary, "M", dlpDate(0).Text, xToDate)
    End If
End If
Call Set_Control("U", Me, rsDATA)
If Val(lblSalaryGrade) = 0 Then rsDATA!SH_GRADE = "00"
rsDATA!sh_compa_user = IIf(optUserSys(0), "", "U")

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    'gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    'gdbAdoIhr001X.CommitTrans
    rsDATA.Resync
    xID = rsDATA("SH_ID")
Else
    'gdbAdoIhr001.BeginTrans
    rsDATA.Update
    'gdbAdoIhr001.CommitTrans
    rsDATA.Requery
    xID = rsDATA("SH_ID")
End If

If xUpdCurrent Then
    Call Set_Current_Flag
End If

Data1.Refresh
DoEvents
prompt = False
Call cmdRecal_Click
DoEvents
prompt = True

Data1.Recordset.Find "SH_ID=" & xID
Data1.Refresh

If glbMediPay Then    'MediPay
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
            If glbCompSerial = "S/N - 2242W" Then
                Call UpdatePTAdministeredBy(glbChgPT, glbChgUseProfile)
            End If
            Call Employee_Transfered_MediPay(glbLEE_ID & "|" & SavPAYP)  ' for #8189
        End If
    End If
End If

glbFlag_BenefitForSalDEPN = False
If OSalary <> medsalary And chkCurrent Then
    Call updBenefitForSalDEPN(glbLEE_ID) 'Jaddy 9/9/99
    If glbCompSerial = "S/N - 2291W" Then Call updCompPlan(glbLEE_ID, Val(medsalary) - Val(OSalary), dlpDate(0).Text)
End If

If chkCurrent Then
    If Not updFollow("U") Then GoTo End_Line 'Exit Sub
End If
'moved to after updFollow by Bryan Ticket#9294
Call Display_Value

DoEvents
If glbGP Then 'George Mar 7,2006 #9965
    Call Salary_Integration(glbLEE_ID, , False, fglbNew, xID) 'George Mar 7,2006 #9965
Else
    Call Salary_Integration(glbLEE_ID)
End If
'medipay doesn't need the employee master tansfer here
Dim saveMedipay
saveMedipay = glbMediPay: glbMediPay = False
Call Employee_Master_Integration(glbLEE_ID)
glbMediPay = saveMedipay

fglbEmptyNew = False
fglbNew = False

glbDisabled = False

Call SET_UP_MODE
Screen.MousePointer = DEFAULT
If glbOttawaCCAC Then
    If chkCurrent Then
        If clpCode(4).Text = "E" Then
            Dim oWHRS, oPHRS
            oWHRS = GetJHData(glbLEE_ID, "JH_WHRS", 0)
            oPHRS = GetJHData(glbLEE_ID, "JH_PHRS", 0)
            If oWHRS = 0 And oPHRS = 0 Then
                MsgBox "Please enter Hours/Week and Hours/Pay Period on Emplopee Position screen."
                Exit Sub
            Else
                If oWHRS = 0 Then
                    MsgBox "Please enter Hours/Week on Emplopee Position screen."
                    Exit Sub
                End If
                If oPHRS = 0 Then
                    MsgBox "Please enter Hours/Pay Period on Emplopee Position screen."
                    Exit Sub
                End If
            End If
        End If
    End If
End If
    
If glbCompSerial = "S/N - 2351W" Then    'Burlington Tech
    'rehiring is using most of the fields in HRAUDIT. Ticket#9899
    rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    If Len(glbChgTermReason) > 0 Then
        Call TermRehireAudit(rsTA)
    End If
    rsTA.Close
End If
If gsEMAIL_ONSALARY Then
    If Len(MailBody) > 0 Then
        If glbFlag_BenefitForSalDEPN Then
            MailBody = MailBody & "The Salary dependent benefits changed too." & vbCrLf
        End If
        Screen.MousePointer = DEFAULT
        Call imgEmail_Click
    End If
End If

Call NextForm
End_Line:
Exit Sub

Add_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdPerform_Click()
'Unload frmEPERFORM
'glbSetPer = glbSetSal
'frmEPERFORM.Show
'Unload Me
'End Sub

Private Sub cmdPerform_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPosition_Click()
'Unload frmEPOSITION
'glbSetPos = glbSetSal
'frmEPOSITION.Show
'Unload Me

'End Sub

Private Sub cmdPosition_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Salary History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgridSal.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HR_SALARY_HISTORY.SH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    xReport = glbIHRREPORTS & "rgridSa2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_SALARY_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Salary History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgridSal.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HR_SALARY_HISTORY.SH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    xReport = glbIHRREPORTS & "rgridSa2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_SALARY_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 0

'cmdPrint.Enabled = True

End Sub



Private Sub CodeEnter(Indx As Integer)

If fglbReason$ <> "NEWH" And Indx < 4 Then
    If Len(clpCode(Indx).Text) > 0 Then
        medPercentChng(Indx).Enabled = True
        medAmtChng(Indx).Enabled = True
    Else
        medPercentChng(Indx) = 0
        medPercentChng(Indx).Enabled = False
        medAmtChng(Indx) = 0
        medAmtChng(Indx).Enabled = False
    End If
End If

End Sub

Private Sub cmdRecal_Click()
Dim xSalary
Dim Msg, a%

If prompt <> False Then
    Msg = "Are You Sure You Want To Recalculate the Percentage and Amount Change(s) For This Employee? "
    Msg = Msg & Chr(10) & Chr(10) & " This Action Will Ignore Records Have Multi-Reason. "
    a% = MsgBox(Msg, 36, "Confirm Recalulate")
    If a% <> 6 Then Exit Sub
End If

Data1.Refresh 'added by Bryan 05-08-05 Ticket #9063
With Data1.Recordset
    If .EOF Then Exit Sub
    xSalary = 0
    .MoveLast
    Do Until .BOF
        If IsNull(.Fields("SH_SREAS2")) And IsNull(.Fields("SH_SREAS3")) Then
            If xSalary = 0 Then
                .Fields("SH_SALPC1") = 1
                .Fields("SH_SALCHG1") = 0
            Else
                .Fields("SH_SALPC1") = (.Fields("SH_SALARY") - xSalary) / xSalary
                .Fields("SH_SALCHG1") = (.Fields("SH_SALARY") - xSalary)
            End If
            .Update
        End If
        xSalary = .Fields("SH_SALARY")
        .MovePrevious
    Loop
    .MoveFirst
End With

Call Set_COMPA
If prompt <> False Then
    DoEvents
    Data1.Refresh
    If Not glbSQL And Not glbOracle Then Call Pause(0.3)
    Data1.Refresh
    DoEvents
    If Not glbSQL And Not glbOracle Then Call Pause(0.3)
    Display_Value
    DoEvents
    Screen.MousePointer = DEFAULT
End If

End Sub

Private Sub cmdRecal_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayPer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayPer_LostFocus()
Dim z%
If comPayPer.ListIndex = 0 Then
    lblSalCode.Caption = "A"
ElseIf comPayPer.ListIndex = 1 Then
    lblSalCode.Caption = "H"
'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
ElseIf glbCompSerial = "S/N - 2282W" And comPayPer.ListIndex = 3 Then
    lblSalCode.Caption = "D"
Else
    lblSalCode.Caption = "M"
End If

If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
z% = getJOB(clpPostCode.Text, clpGrid.Text)
End If
End Sub


Private Sub FIND_JOB()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
On Error GoTo Job_Err
Dim rsJOBs As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT JB_CODE FROM HRJOB"
rsJOBs.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If rsJOBs.EOF And rsJOBs.BOF Then
    Msg = "No Job descriptions found" & Chr(10)
    Msg = Msg & "You will require authority to add one to continue"
    MsgBox Msg
End If
'If Not IsNull(rsJOBs("JB_BAND")) Then
'    fglbBAND = IIf(IsNull(rsJOBs("JB_BAND")), "", rsJOBs("JB_BAND"))
'    lblBANDCode = fglbBAND
'End If

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Jobs", "HRJOB", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next
 
End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo JobHis_Err

Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If

If Not dynaJobHIS.EOF Then
    fglbJobList = ""
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        dynaJobHIS.MoveNext
    Loop
    If Right(fglbJobList, 1) = "," Then
        fglbJobList = Left(fglbJobList, Len(fglbJobList) - 1)
    End If
    dynaJobHIS.MoveFirst
        
End If
Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub
Private Sub Set_NextReview()
Dim EMP_Snap As New ADODB.Recordset
Dim SQLQ, xDATE, xLinDate, NewDate, dtY1%, dtY2%
    'Get Linamar Start Date
    SQLQ = "Select ED_EMPNBR,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    EMP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not (EMP_Snap.BOF And EMP_Snap.EOF) Then
        xLinDate = EMP_Snap("ED_DOH")
        If IsDate(xLinDate) Then
            xDATE = CurSHDate()
            
            If IsDate(xDATE) Then
                dtY1% = DateDiff("yyyy", CVDate(xLinDate), CVDate(xDATE))
                NewDate = DateAdd("yyyy", (dtY1% + 1), CVDate(xLinDate))
            Else
                NewDate = DateAdd("m", 3, CVDate(xLinDate))
            End If
            dlpDate(1) = NewDate
        End If
    End If
    EMP_Snap.Close
    
End Sub
Private Function CurSHDate()
Dim SQLQ As String
Dim HRSH_Snap As New ADODB.Recordset

CurSHDate = 0    ' returns 0 if no found records

On Error GoTo JS_Err

SQLQ = "Select * from HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND SH_CURRENT <>0"
'Town of Aurora or City of Timmins or City of Kawartha Lakes
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2363W" Then
    SQLQ = SQLQ & " ORDER BY SH_SALARY"
ElseIf glbMulti And glbVadim Then
    SQLQ = SQLQ & " AND SH_PAYROLL_ID='" & txtPayrollID.Text & "'"
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
ElseIf glbMulti Then
    SQLQ = SQLQ & " AND SH_JOB='" & clpPostCode.Text & "'"
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
End If
HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRSH_Snap.BOF And HRSH_Snap.EOF Then
    OSalary = 0
    OSalCD = ""
    OJOB = ""
    OEDate = "01/01/01"
    ONDate = "01/01/01"
    OReason = ""
    OLambtonJob = ""
    SavPAYP = ""
    OldPAYP = ""
    oGrade = "00"
    OPremium = "": OTotal = "": OvGroup = "": OVStep = ""
Else
    'Not Town of Aurora and City of Timmins and not City of Kawartha Lakes
    If Not glbCompSerial = "S/N - 2378W" And Not glbCompSerial = "S/N - 2375W" And Not glbCompSerial = "S/N - 2363W" Then
        If fglbNew Then
            If glbMulti And glbVadim Then
                If HRSH_Snap("SH_PAYROLL_ID") = Data1.Recordset("SH_PAYROLL_ID") Then
                    HRSH_Snap("SH_CURRENT") = 0
                    HRSH_Snap.Update
                End If
            End If
        End If
    End If
    CurSHDate = HRSH_Snap("SH_EDATE")
    OSalary = HRSH_Snap("SH_SALARY")
    OSalCD = HRSH_Snap("SH_SALCD")
    OJOB = HRSH_Snap("SH_JOB")
    OEDate = HRSH_Snap("SH_EDATE")
    ONDate = HRSH_Snap("SH_NEXTDAT")
    OReason = HRSH_Snap("SH_SREAS1")
    OLambtonJob = Left(HRSH_Snap("SH_GRID"), 1) & HRSH_Snap("SH_JOB") & Mid(HRSH_Snap("SH_GRID"), 2)
    SavPAYP = HRSH_Snap("SH_PAYP")
    OldPAYP = SavPAYP
    oGrade = HRSH_Snap("SH_GRADE")
    If glbCompSerial = "S/N - 2373W" Then 'Muskoka
        OPremium = HRSH_Snap("SH_PREMIUM"): OTotal = HRSH_Snap("SH_TOTAL")
        OvGroup = HRSH_Snap("SH_VGROUP"): OVStep = HRSH_Snap("SH_VSTEP")
    End If
End If

HRSH_Snap.Close
Exit Function

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SALARY History Snap", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Function

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
    If glbCompSerial = "S/N - 2259W" Then 'Added by Bryan 11/07/05 Ticket #8857 Oxford
        If glbtermopen Then
            SQLQ = "Select ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        Dim rs As New ADODB.Recordset
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_SECTION") = "Y" Then
            glbMulti = True
            lblPayID.Visible = True
            txtPayrollID.Visible = True
        Else
            glbMulti = False
            lblPayID.Visible = False
            txtPayrollID.Visible = False
        End If
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If

If glbtermopen Then
    If glbCompSerial = "S/N - 2191W" Then 'A.E.F.O.
        vbxTrueGrid.Columns(5).NumberFormat = "0.0"
    End If
    If glbOracle Then
        SQLQ = SQLQ & "SELECT Term_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
    Else
        SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
    End If
    
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
Else
    If glbCompSerial = "S/N - 2191W" Then 'A.E.F.O.
        SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
        SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
        SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
        vbxTrueGrid.Columns(5).NumberFormat = "0.0"
    Else
        If glbOracle Then
             SQLQ = SQLQ & "SELECT HR_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        Else
             SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        End If
        SQLQ = SQLQ & "WHERE SH_EMPNBR = " & glbLEE_ID
    End If
End If
SQLQ = SQLQ & " ORDER BY "
If glbMulti Then
    SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_EDATE DESC"
Else
    SQLQ = SQLQ & " SH_EDATE DESC, SH_ID DESC, SH_CURRENT " & IIf(glbSQL, "DESC", "")
End If

If glbCompSerial = "S/N - 2351W" Then   'Burlington Tech.
    vbxTrueGrid.Columns(5).Visible = False
End If

Data1.RecordSource = SQLQ
Data1.Refresh
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    Data1.Recordset.MoveFirst
    Data1.Recordset.Find "SH_CURRENT<>0"
End If
If glbWFC Then
    'Get Employee Plant code
    Call GetPlantCode
End If
EERetrieve = True

Call Display_Value

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Salary", "HR_SALARY_HISTORY", "SELECT")
Unload Me
Resume Next
Exit Function

End Function

Private Sub comSalScale_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub comSalScale_Click()
Dim ssalary, HoursPerWeek!
Dim z%

If fglbGridType = 0.5 And Val(comSalScale) > 0 Then
    lblSalaryGrade = Format((Val(comSalScale) * 2 - 1), "00")
Else
    lblSalaryGrade = Format(Val(comSalScale), "00")
End If

If glbLambton Then 'Ticket# 6693
    If glbSetSal Then
        Exit Sub
    End If
End If

If lblSalaryGrade <> "00" Then
    HoursPerWeek! = Val(lblWhrs)
    
    ssalary = JobSnaps_PayScale(Val(lblSalaryGrade))
    If JobSnaps_Salary_Code$ = "H" Then
        If lblSalCode = "H" Then
            medsalary = Round2DEC(ssalary)
        'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
        ElseIf lblSalCode = "D" Then
            If IsDate(dlpDate(0)) Then
                If GetLeapYear(Year(dlpDate(0))) Then
                    medsalary = Round2DEC((ssalary * HoursPerWeek!) * 366) / 52
                Else
                    medsalary = Round2DEC((ssalary * HoursPerWeek!) * 365) / 52
                End If
            End If
        ElseIf lblSalCode = "A" Then
            If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                medsalary = Round2DEC(ssalary)
            Else
            medsalary = Round2DEC((ssalary * HoursPerWeek!) * 52)
            End If
        ElseIf lblSalCode = "M" Then
            medsalary = Round2DEC(((ssalary * HoursPerWeek!) * 52) / 12)
        End If
    ElseIf JobSnaps_Salary_Code$ = "A" Then
        If lblSalCode = "H" Then
            If HoursPerWeek! = 0 Then
                medsalary = 0
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    medsalary = Round2DEC(ssalary)
                Else
                medsalary = Round2DEC((ssalary / HoursPerWeek!) / 52)
                End If
            End If
        'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
         ElseIf lblSalCode = "D" Then
            If IsDate(dlpDate(0)) Then
                If GetLeapYear(Year(dlpDate(0))) Then
                    medsalary = Round2DEC(ssalary * 366)
                Else
                    medsalary = Round2DEC(ssalary * 365)
                End If
            End If
        ElseIf lblSalCode = "A" Then
            medsalary = Round2DEC(ssalary)
        ElseIf lblSalCode = "M" Then
            medsalary = Round2DEC(ssalary / 12)
        End If
    End If
    medsalary = Round2DEC(Val(medsalary))
    Call setPercent
End If
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMESALARY11"

fglbNew = False
flgloaded = True
glbDisabled = False
Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMESALARY11"
End Sub



Sub Form_Load()
flagFrmLoad = True 'carmen may 00
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
On Error GoTo Err_Deal

fraSalary2.BorderStyle = 0

If glbVadim Then
    lblPayID.FontBold = True
End If
If glbLambton Then
    lblLambtonJob.Visible = True
    txtLambtonJob.Visible = True
End If
If glbMulti Then
    lblPayID.Visible = True
    txtPayrollID.Visible = True
End If
If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If

If glbWFC Then
    dlpDate(2).DataField = "SH_TRANSDATE"
    txtFiscalYear.DataField = "SH_FISCALYEAR"
    clpCode(0).DataField = "SH_SECTION"
    txtMarketLine.DataField = "SH_MARKETLINE"
Else
    dlpDate(2).Enabled = False
End If
'added by Bryan 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    fraSalary.Height = 1515
    medPremium.DataField = "SH_PREMIUM"
    medTotal.DataField = "SH_TOTAL"
    txtVGroup.DataField = "SH_VGROUP"
    txtVStep.DataField = "SH_VSTEP"
Else
    fraSalary.Height = 555
    'fraSalary.Width = 5150
    fraSalary2.Top = 4850
End If
'end bryan

'added by bryan 24/Oct/05 Ticket#9607
If glbCompSerial = "S/N - 2378W" Then 'Aurora
    txtPayrollID.Enabled = False
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

glbOnTop = "FRMESALARY11"

If glbSyndesis Then
    lblTitle(9).Caption = "Range"
    comSalScale.Tag = "00-Posion Grid Ranges"
End If

Call DecSetup

Call FIND_JOB
Call setCaption(lblTitle(12))
Call setCaption(lblGrid)
comPayPer.Clear
comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "
'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
If glbCompSerial = "S/N - 2282W" Then
    comPayPer.AddItem "Daily "
End If
Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
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
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If Len(glbBand) > 0 Then
                If InStr(1, ",A,B,C,D,E,", "," & glbBand & ",") = 0 Then
                    MsgBox "You Do Not Have Authority For This Transaction"
                    glbOnTop = Empty
                    Unload Me
                    Screen.MousePointer = DEFAULT
                    Exit Sub
                End If
            End If
        End If
    End If
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub

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
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If Len(glbBand) > 0 Then
                If InStr(1, ",A,B,C,D,E,", "," & glbBand & ",") = 0 Then
                    MsgBox "You Do Not Have Authority For This Transaction"
                    glbOnTop = Empty
                    Unload Me
                    Screen.MousePointer = DEFAULT
                    Exit Sub
                End If
            End If
        End If
    End If
End If



If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub
If glbWFC Then
    Call Set_COMPA
    Call fgetSection(lblEEID.Caption)
    If fSection = "GREN" Then
        lblTitle(12).FontBold = True
    End If
End If
Screen.MousePointer = HOURGLASS

Call DoWFCGrids(False)
If glbCompSerial = "S/N - 2291W" Then
    lblBANDCode.DataField = ""
    lblBand.Caption = "Mid-Point"
    lblBand.Visible = True
    lblBANDCode.Visible = True
End If
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = IIf(glbSetSal, "Set ", "") & "Salary History- " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbPayWeb Or glbVadim Or glbLambton Or glbInsync Or glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2192W" Then
    lblTitle(12).FontBold = True
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)

lblEEID = glbLEE_ID

Call CR_JobHis_Snap
Call Set_Position(fglbJob$, False)
clpGrid.TABLTitle = lStr(lblGrid)
Call Display_Value

If glbCompSerial = "S/N - 2191W" Then
    fglbFrmt = "0.0"
    lblTitle(12).Caption = "Pay Type"
    clpCode(4).TABLTitle = "Pay Type Codes"
    clpCode(4).Tag = "Enter Pay Type Code"
Else
    fglbFrmt = "00"
End If
If glbOttawaCCAC Or glbCompSerial = "S/N - 2229W" Then  'Ottawa CCAC, Inscape
    lblTitle(12).FontBold = True
End If

If (glbCompSerial = "S/N - 2242W") Then 'C.C.A.C. London & Middlesex - Ticket #6718
    lblTitle(12).FontBold = True
    lblTitle(12).Caption = "Client #"
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    lblComment.Caption = lStr("Union")
    txtComment.Visible = False
    clpCode(5).Left = 1600 '1440
    clpCode(5).Top = 5600 '5520
    clpCode(5).Visible = True
End If
clpGrid.TextBoxWidth = 1000

Dim vList
'Added by Bryan 23/Sep/05 Ticket#9343
cboVGRoup.Clear
cboVStep.Clear
vList = VGroupList
X = 1
Do While X > 0
    X = InStr(vList, "&")
    If X > 0 Then
        cboVGRoup.AddItem Left(vList, X - 1)
        vList = Mid(vList, X + 1)
    Else
        cboVGRoup.AddItem vList
    End If
Loop
vList = VStepList
X = 1
Do While X > 0
    X = InStr(vList, "&")
    If X > 0 Then
        cboVStep.AddItem Left(vList, X - 1)
        vList = Mid(vList, X + 1)
    Else
        cboVStep.AddItem vList
    End If
Loop



Call INI_Controls(Me)
clpGrid.SecurityMaintainable = False
Screen.MousePointer = DEFAULT
Exit Sub

Err_Deal:
If Err = 364 Then Resume Next

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
    Set frmESALARY11 = Nothing
    Call NextForm
End Sub

Private Sub GetPlantCode()
Dim SQLQ As String, xPlantCode
Dim rsXEMP As New ADODB.Recordset
    glbEmpPlant = ""
    SQLQ = "SELECT ED_EMPNBR,ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsXEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsXEMP.EOF Then
        xPlantCode = rsXEMP("ED_SECTION")
    End If
    rsXEMP.Close
    glbEmpPlant = xPlantCode
End Sub
Private Function getJOB(nJob As String, nGrid As String)
Dim SQLQ As String, X%, xLev
Dim Msg$
Dim rsJOB As New ADODB.Recordset
Dim rsDESCR2 As New ADODB.Recordset
'Dim rsGrid As New ADODB.Recordset
'Dim xGridList
getJOB = False
On Error GoTo Jobd_Err
Call setGridList(nJob)
If Len(nJob) > 0 Then
    If glbMultiGrid Then
        SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE = '" & nJob$ & "' AND JB_GRID='" & nGrid & "'"
    Else
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & nJob$ & "' "
    End If
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsJOB.EOF Then
        fglbBAND = ""
        Exit Function
    End If
    If glbCompSerial = "S/N - 2291W" Then
        If Not IsNull(rsJOB("JB_MIDPOINT")) And rsJOB("JB_MIDPOINT") > 0 And rsJOB("JB_MIDPOINT") < 12 Then
            lblBANDCode.Caption = Format(rsJOB("JB_S" & rsJOB("JB_MIDPOINT")), "$0.00")
        Else
            lblBANDCode.Caption = "$0.00"
        End If
    End If
    If glbWFC Then fglbBAND = IIf(IsNull(rsJOB("JB_BAND")), "", rsJOB("JB_BAND"))
    For X% = 1 To 11
        If Not IsNull(rsJOB("JB_S" & X%)) Then JobSnaps_PayScale(X) = Round2DEC(rsJOB("JB_S" & X%))
        
        If glbCompSerial = "S/N - 2378W" And rsJOB("JB_SALCD") <> lblSalCode Then      'Town of Aurora
            If Not IsNull(rsJOB("JB_S" & X% & "A")) Then JobSnaps_PayScale(X) = Round2DEC(rsJOB("JB_S" & X% & "A"))
        End If
    Next
    If Not IsNull(rsJOB("JB_SALCD")) Then JobSnaps_Salary_Code$ = rsJOB("JB_SALCD")
    If Not IsNull(rsJOB("JB_MIDPOINT")) Then JobSnap_MidPoint! = rsJOB("JB_MIDPOINT")
    If Not IsNull(rsJOB("JB_FTEHRS")) Then
        JobSnaps_Salary_FTEHrs = rsJOB("JB_FTEHRS")
    Else
        JobSnaps_Salary_FTEHrs = 1
    End If
    fglbGridType = 0
    
    SQLQ = "SELECT JB_DESCR2,JB_ID FROM HRJOB WHERE JB_CODE='" & nJob & "'"
    rsDESCR2.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsDESCR2.EOF Then
        If IsNumeric(rsDESCR2("JB_DESCR2")) Then
            If Val(rsDESCR2("JB_DESCR2")) = 0.5 Then
                fglbGridType = 0.5
            End If
        End If
    End If
    rsDESCR2.Close
    comSalScale.Clear
    
    comSalScale.AddItem Format(0, fglbFrmt)
    For X% = 1 To 11
        If rsJOB("jb_s" & Trim(str(X%))) <> 0 Then
            xLev = X%
            If fglbGridType = 0.5 Then xLev = (X% + 1) / 2
            
            If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
                If xLev = 1 Then
                    comSalScale.AddItem "Start"
                Else
                    comSalScale.AddItem Format(xLev - 1, fglbFrmt)
                End If
            Else
                comSalScale.AddItem Format(xLev, fglbFrmt)
            End If
        End If
    Next
    
    If fglbGridType = 0.5 And Val(lblSalaryGrade) <> 0 Then
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
            End If
        Else
            comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
        End If
    Else
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
            End If
        Else
            comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
        End If
    End If
    
    If glbWFC Then
        Call Set_MarketLine_List
    End If
End If

getJOB = True

Exit Function

Jobd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "HRJOB", "SELECT")
Resume Next

End Function

Sub Set_MarketLine_List()
Dim rsWFC As New ADODB.Recordset
Dim X%, I%
Dim xItemAdd
Dim SQLQ

SQLQ = "select MarketLine from WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & lblBANDCode & "'"
If Len(clpCode(0)) > 0 Then
    SQLQ = SQLQ & " AND SectionCode ='" & clpCode(0) & "'"
End If
If Len(txtFiscalYear) > 0 Then
    SQLQ = SQLQ & " AND FiscalYear =" & txtFiscalYear & ""
End If
SQLQ = SQLQ & " group by MarketLine"

rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset
X% = 0
cmbMarketLine.Clear
Do Until rsWFC.EOF
    cmbMarketLine.AddItem rsWFC("marketline")
    If rsWFC("marketline") = txtMarketLine Then
        cmbMarketLine.ListIndex = X%
    End If
    X% = X% + 1
    rsWFC.MoveNext
Loop
rsWFC.Close
'MarketLine_Desc Me
Call SalMarketLineDesc

End Sub
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
Private Sub lblBANDCode_Change()
    Set_SalState
End Sub

Private Sub lblCompaNum_Change()
lblCompaNum = Round(Val(lblCompaNum), 2)
End Sub

Private Sub lblSalaryGrade_Change()
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        lblSalaryGrade = Format(Val(lblSalaryGrade), "00")
    End If
    
    If fglbGridType = 0.5 And Val(lblSalaryGrade) > 0 Then
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
            End If
        Else
            comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
        End If
    Else
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
            End If
        Else
            comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
        End If
    End If
End Sub

Private Sub lblSalCode_Change()
If flagFrmLoad = False Then Exit Sub 'carmen may 00
If Len(lblSalCode.Caption) > 0 Then
    If lblSalCode.Caption = "A" Then
        comPayPer.ListIndex = 0
    ElseIf lblSalCode.Caption = "H" Then
        comPayPer.ListIndex = 1
    'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
    ElseIf lblSalCode.Caption = "D" And glbCompSerial = "S/N - 2282W" Then
        comPayPer.ListIndex = 3
    Else
        comPayPer.ListIndex = 2
    End If
End If
End Sub
Sub Set_WFC_COMPA()
Dim xDollear
If glbWFC And UnionExecNone Then
    lblCompaNum = 0
    'If optUserSys(0) Then xDollear = Val(lblsalstate(1)) Else xDollear = Val(mskCampa)
    xDollear = Val(lblsalstate(1))
    'Changed by Bryan 22/Sep/05 Ticket#9343
    
    If Val(xDollear) <> 0 Then
        If glbCompSerial = "S/N - 2373W" Then
            lblCompaNum = (Val(medTotal) / xDollear) * 100
        Else
            lblCompaNum = (Val(medsalary) / xDollear) * 100
        End If
    End If
    If Val(lblCompaNum) > 999.99 Then lblCompaNum = "999.99"
    lblCompaNum.Caption = Format(lblCompaNum, "0.00")
End If
End Sub


Sub Set_SalState()
Dim SQLQ
Dim rsWFC As New ADODB.Recordset
Dim xPlantCd
If Not glbWFC Then Exit Sub
xPlantCd = glbEmpPlant
If Len(clpCode(0).Text) > 0 Then
    xPlantCd = clpCode(0).Text
End If
SQLQ = "SELECT LDOLLARS,MDOLLARS,HDOLLARS FROM WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & Trim(lblBANDCode) & "'"
SQLQ = SQLQ & " AND [MARKETLINE]='" & IIf(txtMarketLine.Visible, txtMarketLine, cmbMarketLine) & "'"
SQLQ = SQLQ & " AND SectionCode='" & xPlantCd & "' "
If Len(txtFiscalYear) > 0 Then
    If IsNumeric(txtFiscalYear) Then
        SQLQ = SQLQ & " AND FiscalYear='" & txtFiscalYear & "' "
    End If
End If

rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenStatic

If rsWFC.EOF Then
  lblsalstate(0) = "": lblsalstate(1) = "": lblsalstate(2) = ""
Else
  lblsalstate(0) = Format(rsWFC("LDOLLARS"), "0.00")
  lblsalstate(1) = Format(rsWFC("MDOLLARS"), "0.00")
  lblsalstate(2) = Format(rsWFC("HDOLLARS"), "0.00")
End If
rsWFC.Close
End Sub


Private Sub medAmtChng_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
'Hemu - essex
'fglbAmtOld(Index) = CCur(Val(medAmtChng(Index)))  'Jaddy 10/25/99
'Hemu - essex
End Sub


Private Sub medAmtChng_KeyPress(Index As Integer, KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medAmtChng_LostFocus(Index As Integer)
If glbSetSal Then Exit Sub
If Not IsNumeric(medAmtChng(Index)) Then
   medAmtChng(Index) = 0
End If

If Not IsNumeric(fglbAmtOld(Index)) Then
   fglbAmtOld(Index) = 0
End If

If medAmtChng(Index) <> fglbAmtOld(Index) Then
    If medAmtChng(Index) <> 0 Then
        If Val(orgSalary) > 0 Then
            medPercentChng(Index) = medAmtChng(Index) / orgSalary
        Else
            medPercentChng(Index) = 1
        End If
    End If
    Call Upd_Salary
End If

Call PerOrSal

End Sub

Private Sub medPercentChng_GotFocus(Index As Integer)

Call SetPanHelp(ActiveControl)

If medPercentChng(Index) = "" Then
   medPercentChng(Index) = 0
End If

medPercentChng(Index) = medPercentChng(Index) * 100
fglbPCOld(Index) = medPercentChng(Index)

End Sub

Private Sub medPercentChng_KeyPress(Index As Integer, KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medPercentChng_LostFocus(Index As Integer)
If Not IsNumeric(medPercentChng(Index)) Then
   medPercentChng(Index) = 0
End If

If Not IsNumeric(fglbPCOld(Index)) Then
   fglbPCOld(Index) = 0
End If
If Not glbSetSal Then
    If medPercentChng(Index) <> fglbPCOld(Index) Then
        ' DK - 03/16/2000 - Removed encryption code
        ' -----
        medAmtChng(Index) = CDbl(medPercentChng(Index)) * orgSalary / 100
        ' -----
        Call Upd_Salary
    End If
End If

medPercentChng(Index) = medPercentChng(Index) / 100
If Not glbSetSal Then
    Call PerOrSal
End If
End Sub

Private Sub medPremium_Change()
Call setPayPeriodSalary
End Sub

Private Sub medPremium_LostFocus()
Dim X%

On Error GoTo Salary_Err 'uncommented 28July99
If Not IsNumeric(medsalary) Then medsalary = 0
medsalary = Round2DEC(Val(medsalary))

If Not IsNumeric(medPremium) Then medPremium = 0
medPremium = Round2DEC(Val(medPremium))

If glbCompSerial = "S/N - 2373W" Then
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    Call Set_SalaryGrade(Val(medTotal))
Else
    Call Set_SalaryGrade(Val(medsalary))
End If
Exit Sub

Salary_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "medPremium", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub

Private Sub medSalary_Change()
    Call setPayPeriodSalary
End Sub
Sub setPayPeriodSalary()
    If glbCompSerial = "S/N - 2373W" Then 'muskoka
        If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
            medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
        ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
            medTotal.Text = medsalary.Text
        End If
        If IsNumeric(medTotal) Then
            'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
            If fglbPhrs <> 0 Then
                If lblSalCode = "H" Then
                    lblPayPeriodSalary = Round2DEC(Val(medTotal) * fglbPhrs)
                    lblHoursPay.Visible = False
                    lblTitle(21).Visible = False
                ElseIf lblSalCode = "M" Then
                    lblPayPeriodSalary = Round2DEC(Val(medTotal))
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    lblHoursPay = Round2DEC(Val(medTotal) / (fglbPhrs * 2))
                'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
                ElseIf lblSalCode = "D" Then
                    If IsDate(dlpDate(0)) Then
                        If GetLeapYear(Year(dlpDate(0))) Then
                            lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 366) / fglbWhrs#) * fglbPhrs)
                        Else
                            lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 365) / fglbWhrs#) * fglbPhrs)
                        End If
                    End If
                ElseIf fglbWhrs# = 0 Then
                    lblPayPeriodSalary = 0
                    lblHoursPay = 0
                Else
                    lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 52) / fglbWhrs#) * fglbPhrs)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    
                    'City of Niagara Falls - Special Hourly Rate calculation
                    If glbCompSerial = "S/N - 2276W" Then
                        lblHoursPay = Round2DEC((Val(medTotal) / fglbPhrs) / (fglbDhrs * 5))
                    Else
                        lblHoursPay = Round2DEC((Val(medTotal) / 52) / fglbWhrs#)
                    End If
                End If
                lblPayPeriodSalary = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                If lblSalCode <> "H" Then
                    lblHoursPay = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                End If
            Else
                lblPayPeriodSalary = 0
                lblHoursPay = 0
            End If
            'Hemu - 08/11/2003 End
        Else
            lblPayPeriodSalary = 0
            lblHoursPay = 0
        End If
    Else
        If IsNumeric(medsalary) Then
            'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
            If fglbPhrs <> 0 Then
                If lblSalCode = "H" Then
                    lblPayPeriodSalary = Round2DEC(medsalary * fglbPhrs)
                    lblHoursPay.Visible = False
                    lblTitle(21).Visible = False
                ElseIf lblSalCode = "M" Then
                    lblPayPeriodSalary = Round2DEC(medsalary)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    lblHoursPay = Round2DEC(Val(medsalary) / (fglbPhrs * 2))
                'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
                ElseIf lblSalCode = "D" Then
                    If IsDate(dlpDate(0)) Then
                        If GetLeapYear(Year(dlpDate(0))) Then
                            lblPayPeriodSalary = Round2DEC(((medsalary / 366) / fglbWhrs#) * fglbPhrs)
                        Else
                            lblPayPeriodSalary = Round2DEC(((medsalary / 365) / fglbWhrs#) * fglbPhrs)
                        End If
                        lblHoursPay.Visible = True
                        lblTitle(21).Visible = True
                        If fglbDhrs <> 0 Then
                            lblHoursPay = Round2DEC(Val(medsalary) / fglbDhrs)
                        Else
                            lblHoursPay = 0
                        End If
                    End If
                ElseIf fglbWhrs# = 0 Then
                    lblPayPeriodSalary = 0
                    lblHoursPay = 0
                Else
                    lblPayPeriodSalary = Round2DEC(((medsalary / 52) / fglbWhrs#) * fglbPhrs)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    'City of Niagara Falls - Special Hourly Rate calculation
                    If glbCompSerial = "S/N - 2276W" Then
                        lblHoursPay = Round2DEC((Val(medsalary) / fglbPhrs) / (fglbDhrs * 5))
                    Else
                        lblHoursPay = Round2DEC((Val(medsalary) / 52) / fglbWhrs#)
                    End If
                End If
                If glbWFC And fSection = "GREN" Then
                    If clpCode(4).Text = "M" Then
                        lblPayPeriodSalary = Round2DEC(medsalary / 12)
                    End If
                End If
                lblPayPeriodSalary = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                If lblSalCode <> "H" Then
                    lblHoursPay = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                End If
            Else
                lblPayPeriodSalary = 0
                lblHoursPay = 0
            End If
        Else
            lblPayPeriodSalary = 0
            lblHoursPay = 0
            'Hemu - 08/11/2003 End
        End If
    End If
End Sub

Private Sub medSalary_GotFocus()

Call SetPanHelp(ActiveControl)
fglbSHold@ = CCur(Val(medsalary))

End Sub

Private Sub medSalary_KeyPress(KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medSalary_LostFocus()
Dim X%

On Error GoTo Salary_Err 'uncommented 28July99
If Not IsNumeric(medsalary) Then medsalary = 0
medsalary = Round2DEC(Val(medsalary))
If Not IsNumeric(medPremium) Then medPremium = 0
medPremium = Round2DEC(Val(medPremium))

If Not glbSetSal Then
    Call setPercent
End If
If glbCompSerial = "S/N - 2373W" Then
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    Call Set_SalaryGrade(Val(medTotal))
Else
    Call Set_SalaryGrade(Val(medsalary))
End If
Exit Sub

Salary_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "medsalary", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub


Private Sub PerOrSal()  'RAUBREY 6/6/97

If Val(medAmtChng(1)) = 0 And Val(medAmtChng(2)) = 0 And Val(medAmtChng(3)) = 0 Then
    fraSalary.Enabled = True
Else
    fraSalary.Enabled = False
End If
End Sub


Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(tmpNUM, glbCompDecHR)
End If

End Function

Private Function Set_Position(nJob As String, nCurrent As Boolean)
Dim SQLQ As String, Msg$
Dim rsHRJob As New ADODB.Recordset

Set_Position = False
On Error GoTo SCError
Screen.MousePointer = HOURGLASS
dynaJobHIS.Requery
dynaJobHIS.Filter = ""
SQLQ = ""
If nCurrent Then SQLQ = SQLQ & " JH_CURRENT<>0 "
If nJob <> "" Then
    SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_JOB='" & nJob & "' "
    'If glbMultiGrid Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_GRID='" & clpGrid.Text & "' "
End If
dynaJobHIS.Filter = SQLQ
Screen.MousePointer = DEFAULT
If dynaJobHIS.BOF And dynaJobHIS.EOF Then
    glbStopSalary% = nCurrent
    Exit Function
Else
    glbStopSalary% = False
End If

If IsNull(dynaJobHIS("JH_WHRS")) Then fglbWhrs# = 0 Else fglbWhrs# = dynaJobHIS("JH_WHRS")
fglbJob$ = dynaJobHIS("JH_JOB")      ' record
fglbSDate = dynaJobHIS("JH_SDATE")
fglbGrid = dynaJobHIS("JH_GRID") & ""
fglbPayrollID = dynaJobHIS("JH_PAYROLL_ID") & ""
orgPosStDate = fglbSDate
If Not IsNull(dynaJobHIS("JH_JREASON")) Then
    fglbReason$ = dynaJobHIS("JH_JREASON")
End If
If Len(dynaJobHIS("JH_ID")) > 0 Then fglbJobID& = dynaJobHIS("JH_ID") Else fglbJobID& = 0

If Len(fglbGrid) > 0 And glbLambton Then txtLambtonJob = Left(fglbGrid, 1) & fglbJob$ & Mid(fglbGrid, 2)
'Hemu
fglbPhrs = dynaJobHIS("JH_PHRS")

'City of Niagara Falls - Pick the Hours per Day from HRJOB table
If glbCompSerial = "S/N - 2276W" Then
    rsHRJob.Open "SELECT JB_CODE, JB_DHRS FROM HRJOB WHERE JB_CODE = '" & dynaJobHIS("JH_JOB") & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRJob.EOF Then
        If IsNull(rsHRJob("JB_DHRS")) Or rsHRJob("JB_DHRS") = "" Then
            fglbDhrs = dynaJobHIS("JH_DHRS")
        Else
            fglbDhrs = rsHRJob("JB_DHRS")
        End If
    Else
        fglbDhrs = dynaJobHIS("JH_DHRS")
    End If
    rsHRJob.Close
Else
    fglbDhrs = dynaJobHIS("JH_DHRS")
End If
'Hemu
dynaJobHIS.Filter = ""
Set_Position = True
Screen.MousePointer = DEFAULT

Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Set_Current_Flag()
Dim SQLQ As String, Msg$
Dim dyn_HRSALHIS As New ADODB.Recordset

On Error GoTo SCFError
If glbMulti Then Exit Sub

'Hemu - 07/07/2003 Begin - Commented out the clone line cause it was giving Error
'                          as 'Row cannot be located for updating'
'Set dyn_HRSALHIS = Data1.Recordset.Clone
dyn_HRSALHIS.Open Data1.RecordSource, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'Hemu- 07/07/2003  End

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

Data1.Refresh

Screen.MousePointer = DEFAULT

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

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

chkCurrent.Enabled = TF
cmdChPos.Enabled = TF
comPayPer.Enabled = TF
comSalScale.Enabled = TF
fraSalary.Enabled = TF
medAmtChng(1).Enabled = TF
medAmtChng(2).Enabled = TF
medAmtChng(3).Enabled = TF
medPercentChng(1).Enabled = TF
medPercentChng(2).Enabled = TF
medPercentChng(3).Enabled = TF
 clpPostCode.Enabled = TF
dlpPosStDate.Enabled = TF
medsalary.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
txtComment.Enabled = TF
cmbMarketLine.Enabled = TF
optUserSys(0).Enabled = TF
optUserSys(1).Enabled = TF
mskCampa.Enabled = TF
If glbSetSal Or glbMulti Then
    clpPostCode.Enabled = TF
    If glbMulti Then
        dlpPosStDate.Enabled = TF
        clpGrid.Enabled = TF
    End If
    cmdChPos.Visible = False
Else
    clpPostCode.Enabled = False
    dlpPosStDate.Enabled = False
    clpGrid.Enabled = False
End If
' danielk - 01/06/2003 - added function to enable editing SH_WHRS for historical records (Ticket #3405)
' danielk - 01/07/2003 - don't enable, only disable in this function, enabling happen w/ edit pos/date btn
If TF = False Then txtWHRS.Enabled = False
' danielk - 01/06/2003 - end
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdRecal.Enabled = False
    cmdChPos.Enabled = False
End If
If glbtermopen Then
    cmdRecal.Enabled = False
'    cmdOK.Enabled = False
'    cmdCancel.Enabled = False
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdChPos.Visible = False
End If
'If Not gSec_Inq_Performance Then cmdPerform.Enabled = False
'If Not gSec_Inq_Position Then cmdPosition.Enabled = False
If glbLinamar Then
    Dim rsTB As New ADODB.Recordset
    rsTB.Open "SELECT ED_EMP FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        If rsTB!ED_EMP = "TEMP" Then
'            cmdNew.Enabled = False
'            cmdModify.Enabled = False
'            cmdDelete.Enabled = False
        End If
    End If
    rsTB.Close
End If

End Sub

Private Sub medTotal_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTotal_LostFocus()
    Call setPayPeriodSalary
End Sub

Sub MskCampa_GotFocus() 'Jaddy 8/9/99
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub mskCampa_LostFocus()
    Call Set_WFC_COMPA
End Sub
Private Sub OptUserSys_Click(Index As Integer) 'Jaddy 8/9/99
End Sub

Private Sub optUserSys_LostFocus(Index As Integer)
    txtUserSys = IIf(optUserSys(0), "", "U")
End Sub

Private Sub optUserSys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 mskCampa.Visible = optUserSys(1)
End Sub


Private Sub txtComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmbMarketLine_GotFocus()   'Jaddy 8/9/99
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub cmbMarketLine_LostFocus()
    txtMarketLine = cmbMarketLine
End Sub

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
Call Set_MarketLine_List
Call Set_SalState
End Sub

Private Sub txtMarketLine_Change() 'Jaddy 8/9/99
  'cmbMarketLine.Clear
  'MarketLine_AddItem Me
  'setMarketLine Me
  Call SalMarketLineDesc
  Call Set_SalState
End Sub

Private Sub txtPosCode_LostFocus()

End Sub

Private Sub Set_COMPA()
Dim SQLQ As String, Msg As String
Dim iRec As Integer
Dim ssalary As Double
Dim X!, cX$
Dim ESalaryCode$
Dim HoursPerWeek!
Dim Compa!
Dim z%
Dim xsSalary  As Double
On Error GoTo UpRel_Err

If glbWFC And UnionExecNone Then Exit Sub

ESalaryCode$ = lblSalCode
'added by Bryan 22/sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    If Len(medTotal) = 0 Then
        ssalary = 0
    Else
        ssalary = medTotal
    End If
Else
    If Len(medsalary) = 0 Then
        ssalary = 0
    Else
        ssalary = medsalary
    End If
End If
HoursPerWeek! = Val(lblWhrs)

If ESalaryCode$ = "H" Then
    If ssalary > 500 Then
        MsgBox "Check if salary is paid Hourly or Annually"
        Exit Sub
    End If
End If
 
z% = getJOB(clpPostCode.Text, clpGrid.Text)
lblBANDCode = fglbBAND
Compa! = 0
If JobSnaps_PayScale(JobSnap_MidPoint!) <> 0 Then

    If JobSnaps_Salary_Code$ = "H" Then
        If ESalaryCode$ = "H" Then
            xsSalary = ssalary
        ElseIf ESalaryCode$ = "M" Then
            If HoursPerWeek! = 0 Then
                xsSalary = 0
            Else
                xsSalary = ((ssalary * 12) / HoursPerWeek!) / 52
            End If
        ElseIf ESalaryCode$ = "A" Then
            If HoursPerWeek! = 0 Then
                xsSalary = 0
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xsSalary = (ssalary)
                Else
                xsSalary = (ssalary / HoursPerWeek!) / 52
                End If
            End If
        End If
    ElseIf JobSnaps_Salary_Code$ = "A" Then
        If ESalaryCode$ = "H" Then
            If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                xsSalary = (ssalary)
            Else
            xsSalary = (ssalary * HoursPerWeek!) * 52
            End If
        ElseIf ESalaryCode$ = "M" Then
            xsSalary = ssalary * 12
        ElseIf ESalaryCode$ = "A" Then
            xsSalary = ssalary
        End If
    End If
    Compa! = (xsSalary / JobSnaps_PayScale(JobSnap_MidPoint!)) * 100
End If
If Compa! > 999.99 Then Compa! = 999.99
If glbCompSerial = "S/N - 2291W" Then Compa! = Round(Compa!, 0) 'Syndesis

lblCompaNum = Compa!

Exit Sub

UpRel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HR_SALARY_HISTORY", "INSERT")
Resume Next

End Sub

Private Sub Upd_Salary()    'RAUBREY 6/6/97
Dim X%
'Hemu - essex
'medSalary = Round2DEC(Val(orgSalary) + CCur(Val(medAmtChng(1))) + CCur(Val(medAmtChng(2))) + CCur(Val(medAmtChng(3))))
medsalary = Round2DEC(Val(orgSalary1) + IIf(fglbAmtOld(1) <> CCur(Val(medAmtChng(1))), CCur(Val(medAmtChng(1))) - fglbAmtOld(1), 0) + IIf(fglbAmtOld(2) <> CCur(Val(medAmtChng(2))), CCur(Val(medAmtChng(2))) - fglbAmtOld(2), 0) + IIf(fglbAmtOld(3) <> CCur(Val(medAmtChng(3))), CCur(Val(medAmtChng(3))) - fglbAmtOld(3), 0))
If fglbAmtOld(1) <> CCur(Val(medAmtChng(1))) Then
    fglbAmtOld(1) = CCur(Val(medAmtChng(1)))
End If
If fglbAmtOld(2) <> CCur(Val(medAmtChng(2))) Then
    fglbAmtOld(2) = CCur(Val(medAmtChng(2)))
End If
If fglbAmtOld(3) <> CCur(Val(medAmtChng(3))) Then
    fglbAmtOld(3) = CCur(Val(medAmtChng(3)))
End If
orgSalary1 = medsalary
If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
    medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
    medTotal.Text = medsalary.Text
End If
'Hemu - essex
' -----

End Sub

Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim Edit1 As Integer

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If fglHredsem <> "" Then    'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'SREV'"
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
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'SREV'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
        ' Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_FREAS") = "SREV"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            Msg = "A Follow Up Record was created!"
            MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpDate(1).Text <> "" Then
        ' 5/2/2001 Add by Frank for no duplicated record of HR_FOLLOW_UP Begin
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'SREV' "
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        

        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
        ' Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_FREAS") = "SREV"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            Msg = "A Follow Up Record was created!"
            MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
  
    If fglbNew = False And Edit1 = True And dlpDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = dlpDate(1).Text
            dynHRAT("EF_FREAS") = "SREV"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglHredsem <> dlpDate(1).Text Then
            Msg = "A Follow Up Record was updated!"
            MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        MsgBox Msg
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
        MsgBox Msg
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
'Private Sub txtPosStDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtPayrollID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPayrollID_KeyPress(KeyAscii As Integer)
If glbVadim Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUserSys_Change()
optUserSys(1) = IIf(txtUserSys = "U", True, False)
optUserSys(0) = Not optUserSys(1)
End Sub

Private Sub txtWHRS_Change()
    lblWhrs.Caption = txtWHRS.Text
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 And Not glbWFC Then
        'dlpDate(2).Text = Updstats(0)
    End If
    If Index = 2 Then
        lblUserDesc = GetUserDesc(Updstats(2))
    End If
End Sub
Private Function GetUserDesc(xUser)
Dim rsUser As New ADODB.Recordset
Dim xDesc
    If Len(xUser) = 0 Then
        xDesc = ""
    Else
        rsUser.Open "SELECT USERID,USERNAME FROM HR_SECURE_BASIC WHERE USERID='" & xUser & "' ", gdbAdoIhr001, adOpenStatic
        If rsUser.EOF Then
            xDesc = xUser
        Else
            xDesc = rsUser("USERNAME")
        End If
        rsUser.Close
    End If
    GetUserDesc = xDesc
End Function
Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

'If KeyAscii = 9 Then ' if the tab key was struck
'    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdClose.SetFocus
'    End If
'End If

End Sub



Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X As Integer, SQLQ

Call Display_Value

fglbJob$ = clpPostCode.Text
Call getJOB(clpPostCode.Text, clpGrid.Text)
optUserSys(1) = IIf(txtUserSys = "U", True, False)
optUserSys(0) = Not optUserSys(1)
mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible
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

Private Sub DecSetup()
If glbCompDecHR = 3 Then
    medsalary.Format = "#,##0.000;(#,##0.000)"
    medTotal.Format = "#,##0.000;(#,##0.000)"
    medPremium.Format = "#,##0.000;(#,##0.000)"
    medAmtChng(1).Format = "#,##0.000;(#,##0.000)"
    medAmtChng(2).Format = "#,##0.000;(#,##0.000)"
    medAmtChng(3).Format = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(1).NumberFormat = "#,##0.000;(#,##0.000)"
End If
If glbCompDecHR = 4 Then
    medsalary.Format = "#,##0.0000;(#,##0.0000)"
    medTotal.Format = "#,##0.0000;(#,##0.0000)"
    medPremium.Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(1).Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(2).Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(3).Format = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(1).NumberFormat = "#,##0.0000;(#,##0.0000)"
End If
End Sub
Private Sub setPercent()
Dim X%
If fglbEmptyNew Then
    medPercentChng(1) = 1
    medAmtChng(1) = medsalary
Else
    If fglbSHold@ <> CCur(medsalary) Then
        For X% = 2 To 3
            medPercentChng(X%) = 0
            medAmtChng(X%) = 0
        Next X%
        medAmtChng(1) = medsalary - orgSalary
        If medAmtChng(1) <> 0 Then
            If orgSalary <> 0 Then
                medPercentChng(1) = medAmtChng(1) / orgSalary
            Else
                medPercentChng(1) = 1
            End If
        Else
            medPercentChng(1) = 0
        End If
    End If
End If
End Sub

Private Sub Get_OrgSalary()
Dim SQLQ As String, HRSH_Snap As New ADODB.Recordset
On Error GoTo JS_Err
SQLQ = "Select SH_SALARY from HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND SH_JOB = '" & clpPostCode.Text & "' "
'Hemu
SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(dlpPosStDate.Text)
'Hemu
SQLQ = SQLQ & " ORDER BY "
If glbMulti Then SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
SQLQ = SQLQ & " SH_EDATE DESC"
HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If HRSH_Snap.BOF And HRSH_Snap.EOF Then
    orgSalary = 0
    orgSalary1 = 0
Else
    orgSalary = HRSH_Snap("SH_SALARY")
    orgSalary1 = HRSH_Snap("SH_SALARY")
End If

HRSH_Snap.Close
Exit Sub

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SALARY History Snap", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Sub

Sub DoWFCGrids(NewEmp As Boolean)
    Dim I As Integer
    
    ' dkostka - 08/31/2000 - WFC requested changes.
    If glbWFC Then
        Data3.ConnectionString = Data1.ConnectionString
        If glbtermopen Then
            Data3.RecordSource = "SELECT ED_ORG,ED_DIV FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            Data3.RecordSource = "SELECT ED_ORG,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
        End If
        Data3.Refresh
        
        If Format(Data3.Recordset("ED_ORG"), "@") = "EXEC" Or Format(Data3.Recordset("ED_ORG"), "@") = "NONE" Then
            UnionExecNone = True
            'If Not NewEmp And Data1.Recordset.EOF = False Then txtMarketLine.DataField = "SH_MARKETLINE"
            'MarketLine_AddItem Me
            
            If NewEmp Then
                If Len(txtMarketLine) = 0 Then 'Ticket# 8046
                    txtMarketLine = GetMarketlineFromDiv(Data3.Recordset("ED_DIV"))
                End If
            End If
            'lblBand.Top = 4500 '4350
            'lblBANDCode.Top = 4500 '4350
            lblBand.Left = 5370
            lblBANDCode.Left = 6600
            lblBand.Visible = True
            lblBANDCode.Visible = True
            lblMarketLine.Visible = True
            cmbMarketLine.Visible = True
            lblMLine.Visible = True
            lblsalstate(0).Visible = True
            lblsalstate(1).Visible = True
            lblsalstate(2).Visible = True
            optUserSys(0).Visible = False 'True Ticket# 6962 WFC doesn't need it
            optUserSys(1).Visible = False 'True Ticket# 6962 WFC doesn't need it
            comSalScale.Visible = False
            lblTitle(9).Visible = False
            lblTitle(13).Visible = True
            lblSalaryGrade.Visible = False
            mskCampa.Visible = False 'True Ticket# 6962 WFC doesn't need it
            lblFiscalYear.Left = 5280 '7200
            txtFiscalYear.Left = 6300
            lblFiscalYear.Visible = True
            txtFiscalYear.Visible = True
            lblTitle(17).Visible = True
            dlpDate(2).Visible = True
            lblPlant.Visible = True
            clpCode(0).Visible = True
            lblPlant.Left = 5280
            clpCode(0).Left = 5980
            fraSalary.Width = 5150
        Else
            UnionExecNone = False
            txtMarketLine.DataField = ""
            comSalScale.Clear
            For I = 1 To 11
                comSalScale.AddItem Format(I, "00")
            Next
            
            lblBand.Visible = False
            lblMarketLine.Visible = False
            cmbMarketLine.Visible = False
            lblMLine.Visible = False
            lblsalstate(0).Visible = False
            lblsalstate(1).Visible = False
            lblsalstate(2).Visible = False
            optUserSys(0).Visible = False
            optUserSys(1).Visible = False
            comSalScale.Visible = True
            lblTitle(9).Visible = True
            lblTitle(13).Visible = False
            lblSalaryGrade.Visible = True
            mskCampa.Visible = False
            lblFiscalYear.Visible = False
            txtFiscalYear.Visible = False
            lblTitle(17).Visible = False
            dlpDate(2).Visible = False
            lblPlant.Visible = False
            clpCode(0).Visible = False
            fraSalary.Width = 9045
        End If
    Else
        ' Not WFC.
        txtMarketLine.DataField = ""
        comSalScale.Clear
        For I = 1 To 11
            If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
                If I = 1 Then
                    comSalScale.AddItem "Start"
                Else
                    comSalScale.AddItem Format(I - 1, "00")
                End If
            Else
                comSalScale.AddItem Format(I, "00")
            End If
        Next
        
        If Not glbSyndesis Then
            lblBand.Visible = False
            lblBANDCode.Visible = False
        End If
        
        lblMarketLine.Visible = False
        cmbMarketLine.Visible = False
        lblMLine.Visible = False
        lblsalstate(0).Visible = False
        lblsalstate(1).Visible = False
        lblsalstate(2).Visible = False
        optUserSys(0).Visible = False
        optUserSys(1).Visible = False
        If (glbCompSerial = "S/N - 2351W") Then 'Burlington Tech
            comSalScale.Visible = False
            lblTitle(9).Visible = False
        Else
            comSalScale.Visible = True
            lblTitle(9).Visible = True
        End If
'        comSalScale.Visible = True
'        lblTitle(9).Visible = True
        lblTitle(13).Visible = False
        lblSalaryGrade.Visible = False
        mskCampa.Visible = False
    End If
    
    ' dkostka end
End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ


'Hemu - 10/09/2003 Begin
Call CR_JobHis_Snap
Call Set_Position(fglbJob$, False)
clpPostCode.seleEMPCode = fglbJobList
'Hemu - 10/09/2003 End

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    lblPayPeriodSalary = ""
    lblHoursPay = ""
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If glbtermopen Then
        If glbCompSerial = "S/N - 2191W" Then
            SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
            SQLQ = SQLQ & " FROM Term_SALARY_HISTORY "
            SQLQ = SQLQ & " LEFT JOIN HRJOB ON Term_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
            SQLQ = SQLQ & " WHERE SH_ID = " & Data1.Recordset!sh_id
            vbxTrueGrid.Columns(5).NumberFormat = "0.0"
        ElseIf glbOracle Then
            SQLQ = SQLQ & "SELECT Term_SALARY_HISTORY.*,SH_GRADE AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
        Else
            SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
        End If
        SQLQ = SQLQ & "WHERE SH_ID = " & Data1.Recordset!sh_id
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        If glbCompSerial = "S/N - 2191W" Then
            SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
            SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
            SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
            vbxTrueGrid.Columns(5).NumberFormat = "0.0"
        ElseIf glbOracle Then
            SQLQ = SQLQ & "SELECT HR_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        Else
            SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        End If
        SQLQ = SQLQ & " WHERE SH_ID = " & Data1.Recordset!sh_id
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If

    If rsDATA.EOF Or rsDATA.BOF Then
    'Hemu - The buttons on the toolbar was not enabling properly if multiple forms
    'were open
    If flgloaded Then
    If UCase(MDIMain.ActiveForm.name) = "FRMESALARY11" Then
    'Hemu
       Call SET_UP_MODE
    End If
    End If
        Exit Sub
    End If
    Call Set_Control("R", Me, rsDATA)
    'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
    Call setPayPeriodSalary
    'Hemu - 08/11/2003 End
    If glbCompSerial = "S/N - 2359W" Then
        clpCode(5) = txtComment
    End If
End If
    If glbLambton Then
        If Len(clpGrid.Text) > 0 And Len(clpPostCode.Text) Then
            txtLambtonJob = Left(clpGrid, 1) & clpPostCode & Mid(clpGrid, 2)
        End If
    End If
    If glbCompSerial = "S/N - 2373W" Then
        If txtVGroup <> "" Then
            cboVGRoup = txtVGroup
        Else
            cboVGRoup = ""
        End If
        If txtVStep <> "" Then
            cboVStep = txtVStep
        Else
            cboVStep = ""
        End If
    End If
    

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
UpdateRight = gSec_Upd_Salary
End Property

Public Property Get Addable() As Boolean

Addable = Not glbtermopen
End Property
Public Property Get Updateble() As Boolean

Updateble = Not glbtermopen
End Property
Public Property Get Deleteble() As Boolean

Deleteble = Not glbtermopen
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
    cmdRecal.Enabled = False
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    cmdRecal.Enabled = False
    TF = False
Else
    UpdateState = OPENING
    TF = True
    cmdRecal.Enabled = True
End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False
Call ST_UPD_MODE(TF)
End Sub


Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmESALARY11.Caption = "Salary - " & Left$(glbLEE_SName, 5)
    frmESALARY11.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
End Sub


Private Sub setGridList(nJob)
    
Dim rsGrid As New ADODB.Recordset
Dim xGridList As String
Dim SaveGrid As String
If Not glbMultiGrid Then Exit Sub
SaveGrid = clpGrid
clpGrid = ""
If Len(clpPostCode.Text) > 0 Then
    rsGrid.Open "SELECT JB_ID,JB_GRID FROM HRJOB_GRADE WHERE JB_CODE='" & CStr(nJob) & "'", gdbAdoIhr001, adOpenForwardOnly
    xGridList = ""
    Do Until rsGrid.EOF
        xGridList = xGridList & "," & rsGrid("JB_GRID")
        rsGrid.MoveNext
    Loop
    If xGridList <> "" Then xGridList = Mid(xGridList, 2)
    clpGrid.seleEMPCode = xGridList
    rsGrid.Close
Else
    clpGrid.seleEMPCode = "NONE-GRID"
End If
clpGrid = SaveGrid
End Sub

Private Sub UpdatePTAdministeredBy(mPT, mAdministeredBy) 'for CCAC London saving Client transfer pop-up window's info
    gdbAdoIhr001.Execute "update HREMP set ED_PT='" & mPT & "', ED_ADMINBY='" & mAdministeredBy & "' where ED_EMPNBR=" & lblEENum
End Sub
Private Function GetMarketlineFromDiv(xDiv)
Dim rsODiv As New ADODB.Recordset
Dim SQLQ, xStr
    xStr = ""
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
    rsODiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsODiv.EOF Then
        If Not IsNull(rsODiv("DV_MARKETLINE")) Then
            xStr = rsODiv("DV_MARKETLINE")
        End If
    End If
    rsODiv.Close
    GetMarketlineFromDiv = xStr
End Function

Private Function VGroupList() As String
Dim retval As String, ctyFile
retval = ""
ctyFile = glbIHRREPORTS & "VGroupList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, retval
    Close #1
End If

ResumeHere:
If InStr(retval, cboVGRoup) = 0 And cboVGRoup <> "" Then
    retval = retval & "&" & cboVGRoup
    cboVGRoup.AddItem cboVGRoup
End If
Open ctyFile For Output As #1
Print #1, retval
Close #1
VGroupList = retval
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt VGroupList.MTF.  INFO:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Function VStepList() As String
Dim retval As String, ctyFile
retval = ""
ctyFile = glbIHRREPORTS & "VStepList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, retval
    Close #1
End If

ResumeHere:
If InStr(retval, cboVStep) = 0 And cboVStep <> "" Then
    retval = retval & "&" & cboVStep
    cboVStep.AddItem cboVStep
End If
Open ctyFile For Output As #1
Print #1, retval
Close #1
VStepList = retval
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt VStepList.MTF.  INFO:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function


Private Sub ResetFlagAudit()
On Error GoTo EH
Dim strSQL As String
Dim rs As New ADODB.Recordset

    strSQL = "SELECT AU_UPLOAD FROM HRAUDIT WHERE AU_EmpNBR=" & glbLEE_ID
    strSQL = strSQL & " AND AU_LDATE=" & Date_SQL(dlpDate(0).Text)
    strSQL = strSQL & " AND AU_SREASON = '" & clpCode(1).Text & "'"
    strSQL = strSQL & " AND AU_SALARY = " & medsalary
    rs.Open strSQL, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic, adCmdText

    If rs.EOF = False And rs.BOF = False Then
        rs("AU_UPLOAD") = "Y"
        rs.Update
    End If
    rs.Close
    
exH:
    Set rs = Nothing
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Updating AUDIT RECORD", "AUDIT FILE", "UPDATE")
    Call RollBack '28July99 js
    Resume exH

    
End Sub
Private Sub ChangeEDateAudit(xEDate)
On Error GoTo EH
Dim strSQL As String
Dim rs As New ADODB.Recordset

    strSQL = "SELECT AU_LDATE, AU_SEDATE FROM HRAUDIT WHERE AU_EMPNBR=" & glbLEE_ID
    strSQL = strSQL & " AND AU_LDATE=" & Date_SQL(xEDate)
    strSQL = strSQL & " AND AU_SREASON = '" & clpCode(1).Text & "'"
    strSQL = strSQL & " AND AU_SALARY = " & medsalary
    rs.Open strSQL, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic, adCmdText

    If rs.EOF = False And rs.BOF = False Then
        rs("AU_LDATE") = dlpDate(0).Text
        rs("AU_SEDATE") = dlpDate(0).Text
        rs.Update
    End If
    rs.Close
    
exH:
    Set rs = Nothing
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Updating AUDIT RECORD", "AUDIT FILE", "UPDATE")
    Call RollBack '28July99 js
    Resume exH

    
End Sub
Private Function fgetSection(xID)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If glbtermopen Then
        strSQL = "SELECT ED_SECTION FROM TERM_HREMP WHERE TERM_SEQ =" & xID
        rs.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
    Else
        strSQL = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & xID
        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    End If
    
    If rs.EOF = False Then
        If Not IsNull(rs("ED_SECTION")) Then
            fSection = rs("ED_SECTION")
        Else
            fSection = ""
        End If
    End If
    rs.Close
    Set rs = Nothing
    

End Function

Public Sub imgEmail_Click()
Dim xEmail
On Error GoTo Email_Err
    If gsEMAIL_ONSALARY Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = GetComPreferEmail("EMAIL_ONSALARY")
            frmSendEmail.txtCC.Text = xEmail
            frmSendEmail.txtSubject.Text = "INFO:HR Salary Change Notice"
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        Else
            If Len(glbLEE_SName) = 0 Then
                MsgBox "There is no email on Status/Dates screen for employee. "
            Else
                MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            End If
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


