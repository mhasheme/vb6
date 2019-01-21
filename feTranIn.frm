VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmETRANIN 
   Appearance      =   0  'Flat
   Caption         =   "Transfer In"
   ClientHeight    =   10950
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   19440
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   19440
   Tag             =   "01-Employee ID in the Division"
   WindowState     =   2  'Maximized
   Begin VB.Frame frmLinLabourCode 
      Height          =   330
      Left            =   8760
      TabIndex        =   136
      Top             =   5880
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtLabCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "JH_LabourCD"
         Height          =   285
         Left            =   320
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "00-Bonus Reporting #"
         Top             =   0
         Width           =   990
      End
      Begin VB.Image imgILabCode 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   0
         Picture         =   "feTranIn.frx":0000
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblLabCodeDesc 
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   137
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdEditPayID 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      Height          =   280
      Left            =   1320
      TabIndex        =   100
      Tag             =   "Edit Transaction Date"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   13
      Left            =   12960
      TabIndex        =   35
      Tag             =   "00-Shift"
      Top             =   3480
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SHFT"
      MaxLength       =   8
   End
   Begin VB.TextBox txtUserNum1 
      Appearance      =   0  'Flat
      DataSource      =   " "
      Height          =   285
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   27
      Tag             =   "00-User Number 1"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtUserText2 
      Appearance      =   0  'Flat
      DataField       =   "ED_USER_TEXT2"
      DataSource      =   " "
      Height          =   285
      Left            =   4425
      MaxLength       =   20
      TabIndex        =   28
      Tag             =   "00-User Text 2"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ComboBox comUserText2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Tag             =   "00-User Text 2"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CheckBox chkProSha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13560
      TabIndex        =   52
      Tag             =   "40-Lead Hand - y/n"
      Top             =   5190
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtEmpType 
      Appearance      =   0  'Flat
      DataSource      =   " "
      Height          =   285
      Left            =   12960
      MaxLength       =   15
      TabIndex        =   110
      Tag             =   "00-Internal Telephone Extension "
      Top             =   7200
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ComboBox comEmpType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "feTranIn.frx":014A
      Left            =   13680
      List            =   "feTranIn.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "10-Type of Employee "
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtFiscalYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7820
      MaxLength       =   4
      TabIndex        =   42
      Tag             =   "00-Fiscal Year"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbMarketLine 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7820
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Tag             =   "00-Market Line"
      Top             =   6945
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtMarketLine 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "SH_MarketLine"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10200
      TabIndex        =   102
      Top             =   6600
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.TextBox txtPayrollID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2120
      MaxLength       =   25
      TabIndex        =   6
      Tag             =   "01-Employee Payroll ID"
      Top             =   2400
      Width           =   2505
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   2
      Left            =   7500
      TabIndex        =   38
      Tag             =   "Next Review Date"
      Top             =   5280
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   7500
      TabIndex        =   30
      Tag             =   "01-Position code"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpGLNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Tag             =   "G/L Number"
      Top             =   4560
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "Department"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "01-Division"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2250
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Reason for Transfer"
      Top             =   930
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDJC"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   315
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Tag             =   "Facility Start Date"
      Top             =   540
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7820
      MaxLength       =   1
      TabIndex        =   34
      Tag             =   "00-Code assigned to the shift"
      Top             =   5610
      Width           =   870
   End
   Begin VB.ComboBox comPayPer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Tag             =   "01-Choose annum or hour"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtEmpID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2120
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "01-Employee ID in the Division"
      Top             =   2040
      Width           =   825
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   60
      Top             =   10290
      Width           =   19440
      _Version        =   65536
      _ExtentX        =   34290
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
      Begin VB.CommandButton cmdFrankTest 
         Height          =   375
         Left            =   10920
         TabIndex        =   135
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Left            =   1335
         TabIndex        =   90
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
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
         Left            =   480
         TabIndex        =   89
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "JH_DHRS"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   7830
      TabIndex        =   31
      Tag             =   "10-Usual working hours per day"
      Top             =   3480
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
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   7830
      TabIndex        =   32
      Tag             =   "10- Number of hours in work week"
      Top             =   3840
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
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   7830
      TabIndex        =   33
      Tag             =   "10-Usual working hours per pay period"
      Top             =   4200
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
   Begin MSMask.MaskEdBox medSalary 
      DataField       =   "SH_SALARY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7830
      TabIndex        =   36
      Tag             =   "21-Enter salary"
      Top             =   4560
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   19440
      _Version        =   65536
      _ExtentX        =   34290
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
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
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
         Left            =   9720
         TabIndex        =   57
         Top             =   4080
         Width           =   945
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
         TabIndex        =   55
         Top             =   135
         Width           =   1185
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
         Left            =   1200
         TabIndex        =   56
         Top             =   135
         Width           =   1245
      End
   End
   Begin MSMask.MaskEdBox medSIN 
      Height          =   285
      Left            =   2130
      TabIndex        =   16
      Tag             =   "00-Social Insurance Number"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###-###-###"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   1
      Left            =   8250
      TabIndex        =   1
      Tag             =   "Date Transferred In"
      Top             =   570
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Tag             =   "Product Line"
      Top             =   3480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Tag             =   "Home Operation Number"
      Top             =   3840
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMOP"
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Tag             =   "Home Line"
      Top             =   4200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMLN"
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   13
      Tag             =   "Home Work Center"
      Top             =   4920
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMWC"
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   14
      Tag             =   "Home Shift"
      Top             =   5280
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMSF"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Tag             =   "Operation"
      Top             =   5640
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   7500
      TabIndex        =   39
      Tag             =   "Labour Code"
      Top             =   5940
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDLB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   18
      Tag             =   "00-Administered By"
      Top             =   8010
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   19
      Top             =   8400
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   17
      Tag             =   "00-Location - Code"
      Top             =   7440
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   7500
      TabIndex        =   41
      Tag             =   "01-Reason for change in position - Code"
      Top             =   6270
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataSource      =   " "
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   20
      Tag             =   "41-Original Hire Date "
      Top             =   8760
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1060
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataSource      =   " "
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   21
      Tag             =   "40-Seniority Date"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1060
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataSource      =   " "
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   22
      Tag             =   "40-Last Hire Date"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1060
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   7500
      TabIndex        =   44
      Tag             =   "00-Enter pay period code"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin INFOHR_Controls.CodeLookup clpVadim2 
      Height          =   285
      Left            =   7500
      TabIndex        =   45
      Top             =   7680
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDV2"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   " "
      Height          =   285
      Index           =   10
      Left            =   13395
      TabIndex        =   24
      Tag             =   "00-Enter Union Code"
      Top             =   6840
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpBGroup 
      Height          =   285
      Left            =   13395
      TabIndex        =   26
      Tag             =   "01-Benefit - Group Code"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BGMF"
      MaxLength       =   10
      SecurityMaintainable=   0
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   6
      Left            =   8250
      TabIndex        =   3
      Tag             =   "Other Date 1"
      Top             =   960
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpGrid 
      Height          =   285
      Left            =   7500
      TabIndex        =   47
      Top             =   8010
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGD"
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      Height          =   285
      Index           =   2
      Left            =   13050
      TabIndex        =   50
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   4500
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      Height          =   285
      Index           =   1
      Left            =   13050
      TabIndex        =   49
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   4170
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      Height          =   285
      Index           =   0
      Left            =   13050
      TabIndex        =   48
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      Height          =   285
      Index           =   3
      Left            =   13050
      TabIndex        =   51
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   4845
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpSalDist 
      Height          =   285
      Left            =   13440
      TabIndex        =   121
      Top             =   8280
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   11
      Left            =   13440
      TabIndex        =   123
      Tag             =   "00-Supervisory Code for cheque sorting "
      Top             =   7920
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSP"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataSource      =   " "
      Height          =   285
      Index           =   7
      Left            =   13440
      TabIndex        =   125
      Tag             =   "40-OMERS Date"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1105
   End
   Begin INFOHR_Controls.CodeLookup clpVadim1 
      Height          =   285
      Left            =   13320
      TabIndex        =   127
      Top             =   9000
      Visible         =   0   'False
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDV1"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataSource      =   " "
      Height          =   285
      Index           =   8
      Left            =   13320
      TabIndex        =   129
      Tag             =   "40-User Defined"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1105
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   12
      Left            =   13320
      TabIndex        =   25
      Tag             =   "00-Enter Status Code"
      Top             =   9720
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   14
      Left            =   13320
      TabIndex        =   46
      Tag             =   "00-Orgranization - Code"
      Top             =   9960
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ORGN"
   End
   Begin VB.Label lblWFCNote5 
      Caption         =   "Don't forget to check for data alignment in the Demographic, Status/Dates, Position and Salary screens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5820
      TabIndex        =   139
      Top             =   8400
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   11160
      TabIndex        =   138
      Top             =   10005
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblWFCNote 
      Caption         =   $"feTranIn.frx":014E
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   13560
      TabIndex        =   134
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblEEStatus2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   133
      Top             =   9720
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblUserNum1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Number 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   132
      Top             =   9885
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblUserText2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Text 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3360
      TabIndex        =   131
      Top             =   9885
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblUDay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Defined"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   130
      Top             =   9360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblVadim11 
      AutoSize        =   -1  'True
      Caption         =   "Vadim Field 1"
      Height          =   195
      Left            =   11160
      TabIndex        =   128
      Top             =   9000
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblODate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OMERS Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   126
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblSupervisor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   124
      Top             =   7920
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label lblSalDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Distribution"
      Height          =   195
      Left            =   11160
      TabIndex        =   122
      Top             =   8280
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Eigible for Profit Sharing"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   11280
      TabIndex        =   120
      Top             =   5190
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblReptAuth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 3"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   11280
      TabIndex        =   119
      Top             =   4500
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblReptAuth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   11280
      TabIndex        =   118
      Top             =   4170
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblReptAuth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   11280
      TabIndex        =   117
      Top             =   3840
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblReptAuth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 4"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   11280
      TabIndex        =   116
      Top             =   4845
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblBand 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Band"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9120
      TabIndex        =   115
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "Grid Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5820
      TabIndex        =   114
      Top             =   8040
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Date 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   6240
      TabIndex        =   113
      Top             =   960
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblNGSStart 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"feTranIn.frx":01F8
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   6960
      TabIndex        =   112
      Top             =   1320
      Visible         =   0   'False
      Width           =   4275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBen 
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Group"
      Height          =   255
      Left            =   11160
      TabIndex        =   111
      Top             =   7560
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   109
      Top             =   6885
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblEEType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11160
      TabIndex        =   108
      Top             =   7245
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblVadim2 
      AutoSize        =   -1  'True
      Caption         =   "Vadim Field 2"
      Height          =   195
      Left            =   5820
      TabIndex        =   107
      Top             =   7680
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   5820
      TabIndex        =   106
      Top             =   7350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblFiscalYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
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
      Left            =   5820
      TabIndex        =   105
      Top             =   6645
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblMarketLine 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Market Line"
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
      Left            =   5820
      TabIndex        =   104
      Top             =   7005
      Visible         =   0   'False
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
      Left            =   9060
      TabIndex        =   103
      Top             =   7005
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblPayIDExist 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Number Already Exist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4680
      TabIndex        =   101
      Top             =   2400
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll ID"
      Height          =   195
      Index           =   26
      Left            =   360
      TabIndex        =   99
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lblLHire 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Hire"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   98
      Top             =   9480
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblSen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seniority"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   97
      Top             =   9120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblOHire 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Original Hire"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   96
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEEStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Change"
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
      Left            =   5820
      TabIndex        =   95
      Top             =   6270
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   360
      TabIndex        =   94
      Top             =   7470
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   360
      TabIndex        =   93
      Top             =   8400
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   360
      TabIndex        =   92
      Top             =   8040
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   91
      Top             =   7080
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblCountry 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country will Changed"
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
      Left            =   360
      TabIndex        =   88
      Top             =   6360
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblPROV 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Province will changed"
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
      Left            =   360
      TabIndex        =   87
      Top             =   6720
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblSIN 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter S.I.N."
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
      Left            =   360
      TabIndex        =   86
      Top             =   6000
      Visible         =   0   'False
      Width           =   1020
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5820
      TabIndex        =   85
      Top             =   5610
      Width           =   405
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Labour Code"
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
      Height          =   255
      Index           =   27
      Left            =   5760
      TabIndex        =   84
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Per"
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
      Index           =   22
      Left            =   5820
      TabIndex        =   83
      Top             =   4920
      Width           =   300
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9060
      TabIndex        =   82
      Top             =   4950
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblEmpNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblEENum"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3090
      TabIndex        =   81
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   21
      Left            =   360
      TabIndex        =   80
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   360
      TabIndex        =   79
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours/Day"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   5820
      TabIndex        =   78
      Top             =   3480
      Width           =   780
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Index           =   19
      Left            =   5820
      TabIndex        =   77
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours/Week"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   5820
      TabIndex        =   76
      Top             =   3840
      Width           =   930
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours/Pay Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   5820
      TabIndex        =   75
      Top             =   4200
      Width           =   1260
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
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
      Left            =   5820
      TabIndex        =   74
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Review Date"
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
      Left            =   5820
      TabIndex        =   73
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Employee Position/Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5580
      TabIndex        =   72
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Work Center"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   71
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   70
      Top             =   4560
      Width           =   435
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Line"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   69
      Tag             =   "Home Line"
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   68
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Basic Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   67
      Top             =   2760
      Width           =   2355
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Operation #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   66
      Top             =   3840
      Width           =   1305
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   360
      TabIndex        =   65
      Top             =   5640
      Width           =   690
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Index           =   6
      Left            =   360
      TabIndex        =   64
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Line "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   63
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Transferred In"
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
      Index           =   4
      Left            =   6480
      TabIndex        =   62
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Transfer"
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
      Index           =   3
      Left            =   360
      TabIndex        =   61
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label lblEmpExist 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number Already Exist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4140
      TabIndex        =   59
      Top             =   2040
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Employee Number"
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
      Index           =   2
      Left            =   120
      TabIndex        =   58
      Top             =   1320
      Width           =   1965
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Division Start Date"
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
      Left            =   360
      TabIndex        =   53
      Top             =   600
      Width           =   1620
   End
End
Attribute VB_Name = "frmETRANIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsORG As New ADODB.Recordset
Dim RDept, RGLNum
Dim fglbTERM_Seq
Dim fglbJobID&, fglbBAND As String
Dim JobSnap_PayScale(20) As Double '15 -> 20
Dim JobSnap_Salary_Code$
Dim JobSnap_MidPoint!
Dim OHOMELINE, OHOMESHIFT, OHOMEOPRTNBR, oHOMEWRKCNT
Dim xUpdateable, ODateOfHire, OldDiv
Dim xWFC_OldNGSStart, xWFC_NewNGSStart
Dim OldUnion, OldDept, OldSection, OldRegion, OldAdminBy, OldLoc
Dim SaveBGroup
Dim xCovClass As String, xBenAccount As String
Dim locSurName, locFName
Dim xTL_OLDEMPNBR 'Ticket #24552 Franks 11/01/2013
Dim xWFCPosChgEmailBody 'Ticket #29343 Franks 10/26/2016
Dim xIsWFCPosChgEmail As Boolean  'Ticket #29343 Franks 10/26/2016
Dim xLocSIN
Dim xPayIDEnable As Boolean

Private Function ChkInSamuel()
Dim X%
Dim Msg$, Response%
Dim xDivCountry As String
ChkInSamuel = False

If Len(dlpDate(0)) > 0 Then
    If Not IsDate(dlpDate(0)) Then
         MsgBox "Invalid Date of Start"
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Start Date is a required field."
    dlpDate(0).SetFocus
    Exit Function
End If
If Len(dlpDate(1)) > 0 Then
    If Not IsDate(dlpDate(0)) Then
        MsgBox "Invalid Date of Transfer In"
        dlpDate(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Date of Transfer In is a required field."
    dlpDate(1).SetFocus
    Exit Function
End If

If Len(clpCode(5)) = 0 Then
    MsgBox "PLANT is a required field"
     clpCode(5).SetFocus
    Exit Function
Else
    If clpCode(5).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
         clpCode(5).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(1)) = 0 Then
    MsgBox "Reason for Transfer In is a required field"
     clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
         clpCode(1).SetFocus
        Exit Function
    End If
End If

If Len(clpDept) = 0 Then
    MsgBox "Department is a required field"
     clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        MsgBox "Invalid Department"
         clpDept.SetFocus
        Exit Function
    End If
End If
If clpDIV.Caption = "Unassigned" Then
    MsgBox lStr("Invalid Division")
    clpDIV.SetFocus
    Exit Function
End If
If Len(clpCode(0)) = 0 Then
    MsgBox lStr("Section is a required field")
    clpCode(0).SetFocus
    Exit Function
Else
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Section")
         clpCode(0).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(7)) = 0 Then
    MsgBox lStr("Location is a required field")
    clpCode(7).SetFocus
    Exit Function
Else
    If clpCode(7).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Location")
         clpCode(7).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(6)) = 0 Then
    MsgBox lStr("Region is a required field")
    clpCode(6).SetFocus
    Exit Function
Else
    If clpCode(6).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Region")
         clpCode(6).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(10)) = 0 Then
    MsgBox lStr("Union is a required field")
    clpCode(10).SetFocus
    Exit Function
Else
    If clpCode(10).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Union")
         clpCode(10).SetFocus
        Exit Function
    End If
End If

'End If
If Len(txtEmpID) = 0 Then
    MsgBox "Employee ID is a required field"
    txtEmpID.SetFocus
    Exit Function
Else
    If Not IsNumeric(txtEmpID) Then
        MsgBox "Invalid Employee ID"
        txtEmpID.SetFocus
        Exit Function
    Else
        If Val(txtEmpID) = 0 Then
            MsgBox "Invalid Employee ID"
            txtEmpID.SetFocus
            Exit Function
        End If
    End If
End If

If lblEmpExist.Visible = True Then
  MsgBox "The NEW Employee number already exits"
  'txtEmpID.SetFocus
  Exit Function
End If

'Ticket #22260 Franks 07/27/2012 - begin
'If "Supervisor" is entered, OMERS Date must be entered too
'If Len(clpCode(11).Text) > 0 Then
'Ticket #22262 Franks 085/02/2012 - pension code field should be Salary Distribution not Supervisor Code
If Len(clpSalDist.Text) > 0 Then
    If Len(dlpDate(7).Text) = 0 Then
        MsgBox lStr("OMERS Date") & " is required if " & lStr("Salary Distribution") & " is entered."
        dlpDate(7).SetFocus
        Exit Function
    End If
End If
If Len(clpVadim1.Text) > 0 Then
    If Len(dlpDate(8).Text) = 0 Then
        MsgBox lStr("User Defined") & " is required if " & lStr("Vadim Field 1") & " is entered."
        dlpDate(8).SetFocus
        Exit Function
    End If
End If
'Ticket #22260 Franks 07/27/2012 - end
    


For X% = 2 To 3
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
     clpCode(X%).SetFocus
    Exit Function
End If
Next X%

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
If Len(medSalary) < 1 Then
    MsgBox "Salary is required"
    medSalary.SetFocus
    Exit Function
End If
If Val(medSalary) <= 0 Then
    MsgBox "Salary is required"
    medSalary.SetFocus
    Exit Function
End If
If comPayPer.ListIndex = -1 Or lblSalCode = "" Then
    MsgBox "Per is required field"
    comPayPer.SetFocus
    Exit Function
End If
If Len(dlpDate(2)) > 0 Then
    If Not IsDate(dlpDate(2)) Then
        MsgBox "Invalid Date "
        dlpDate(2).SetFocus
        Exit Function
    End If
Else
        MsgBox "Effective Date is required"
        dlpDate(2).SetFocus
        Exit Function
End If
If Len(clpCode(8)) = 0 Then
    MsgBox lStr("Reason for Change is a required field")
    clpCode(8).SetFocus
    Exit Function
Else
    If clpCode(8).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Reason for Change")
         clpCode(8).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(9)) = 0 Then
    MsgBox lStr("Pay Period is a required field")
    clpCode(9).SetFocus
    Exit Function
Else
    If clpCode(9).Caption = "Unassigned" Then
        MsgBox lStr("Invalid Pay Period")
         clpCode(9).SetFocus
        Exit Function
    End If
End If

If Len(clpGrid) = 0 Then
    MsgBox lStr("Grid Category is a required field")
    clpGrid.SetFocus
    Exit Function
Else
    If clpGrid.Caption = "Unassigned" Then
        MsgBox lStr("Invalid Grid Category")
        clpGrid.SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(3)) > 0 Then
    If Not IsDate(dlpDate(3)) Then
        MsgBox "Invalid Date"
        dlpDate(3).SetFocus
        Exit Function
    End If
Else
    MsgBox lStr("Original Hire Date is a required field.")
    dlpDate(3).SetFocus
    Exit Function
End If
If Len(dlpDate(4)) > 0 Then
    If Not IsDate(dlpDate(4)) Then
        MsgBox "Invalid Date"
        dlpDate(4).SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(5)) > 0 Then
    If Not IsDate(dlpDate(5)) Then
        MsgBox "Invalid Date"
        dlpDate(5).SetFocus
        Exit Function
    End If
End If

'Ticket #22157 Franks 06/29/2012
If clpCode(5).Text = "5322" Or clpCode(5).Text = "2158" Then
    If Len(clpGLNum.Text) = 0 Then
        MsgBox lStr("G/L #") & " is a required field if " & lStr("Administered By") & " is '5322' or '2158'."
        clpGLNum.SetFocus
        Exit Function
    End If
End If

ChkInSamuel = True


End Function
Private Function ChkInput()
Dim X%
Dim Msg$, Response%
Dim xDivCountry As String
ChkInput = False

If glbLinamar Then 'Ticket #29759 Franks 02/22/2017 check if Payroll ID is duplicate
    If Len(txtPayrollID.Text) > 0 Then
        If IsLinDupPayrollID(glbTran_ID, txtPayrollID.Text, "N", "N", xLocSIN) Then
            MsgBox "Duplicate Payroll ID."
            Exit Function
        End If
    End If
End If

If Len(dlpDate(0)) > 0 Then
    If Not IsDate(dlpDate(0)) Then
        If glbLinamar Then
            MsgBox "Invalid Date of Facility Start"
        End If
        If glbWFC Then
            MsgBox "Invalid Date of Start"
        End If
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    If glbLinamar Then
        MsgBox "Facility Start Date is a required field."
    End If
    If glbWFC Then
        MsgBox "Start Date is a required field."
    End If
    dlpDate(0).SetFocus
    Exit Function
End If
If Len(dlpDate(1)) > 0 Then
    If Not IsDate(dlpDate(0)) Then
        MsgBox "Invalid Date of Transfer In"
        dlpDate(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Date of Transfer In is a required field."
    dlpDate(1).SetFocus
    Exit Function
End If

If Len(clpCode(1)) = 0 Then
    MsgBox "Reason for Transfer In is a required field"
     clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
         clpCode(1).SetFocus
        Exit Function
    End If
End If

If glbLinamar Then
    If clpDIV.Caption = "Unassigned" Or Len(clpDIV) <> 3 Or Not IsNumeric(clpDIV) Then
        MsgBox "Invalid Facility"
         clpDIV.SetFocus
        Exit Function
    End If
    If Len(clpCode(2)) = 0 Then
        MsgBox lStr("Region Code is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbWFC Then
    If clpDIV.Caption = "Unassigned" Or Not IsNumeric(clpDIV) Then
        MsgBox lStr("Invalid Division")
         clpDIV.SetFocus
        Exit Function
    End If
    If Len(clpCode(0)) = 0 Then
        MsgBox lStr("Section is a required field")
        clpCode(0).SetFocus
        Exit Function
    Else
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox lStr("Invalid Section")
             clpCode(0).SetFocus
            Exit Function
        End If
    End If
    
    If Len(clpGLNum.Text) = 0 Then 'Ticket #24317 Franks 09/17/2013
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    
    'Ticket #23247 Franks 07/23/2013
    If clpCode(12).Visible Then
        If Len(clpCode(12)) = 0 Then
            MsgBox lStr("Employment Status") & " is a required field"
            clpCode(12).SetFocus
            Exit Function
        Else
            If clpCode(12).Caption = "Unassigned" Then
                MsgBox lStr("Invalid Employment Status")
                clpCode(12).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Len(clpCode(10).Text) < 1 Then 'Ticket #24317 Franks 09/17/2013
        MsgBox lStr("Union code is a required field")
        clpCode(10).SetFocus
        Exit Function
    End If
    
    'Ticket #16749
    xDivCountry = GetCountryFromDiv(clpDIV.Text)
    If xDivCountry = "U.S.A." Then
        If Len(clpVadim2.Text) = 0 Then
            MsgBox lStr("Vadim Field 2") & " is a required field if this is an US Division"
            clpVadim2.SetFocus
            Exit Function
        End If
    End If
    If lblBen.FontBold = True Then 'Ticket #18654
        If Len(clpBGroup.Text) = 0 Then
            MsgBox "Benefit Group is required since the country is 'CANADA' for " & lStr("Division") & " " & clpDIV.Text
            clpBGroup.SetFocus
            Exit Function
        Else
            If clpBGroup.Caption = "Unassigned" Then
                MsgBox ("Invalid Benefit Group")
                clpBGroup.SetFocus
                Exit Function
            End If
        End If
    End If
    'Ticket #22448 Franks 10/31/2012 - begin
    If lblUserNum1.FontBold = True Then
        If Len(txtUserNum1.Text) = 0 Then
            MsgBox lStr(lblUserNum1.Caption) & " is required since the country is 'CANADA' for " & lStr("Division") & " " & clpDIV.Text
            txtUserNum1.SetFocus
            Exit Function
        Else
            If Not IsNumeric(txtUserNum1.Text) Then
                MsgBox lStr(lblUserNum1.Caption) & " must be a number"
                txtUserNum1.SetFocus
                Exit Function
            End If
        End If
    End If
    If lblUserText2.FontBold = True Then
        If Len(comUserText2.Text) = 0 Then
            MsgBox lStr(lblUserText2.Caption) & " is required since the country is 'CANADA' for " & lStr("Division") & " " & clpDIV.Text
            comUserText2.SetFocus
            Exit Function
        End If
    End If
    'Ticket #22448 Franks 10/31/2012 - end
    
    If glbWFC And glbCandidate > 0 Then 'Ticket #24184 Franks 11/12/2013
        If isValidWFCEmpNo(glbLEE_ID, clpDIV.Text, clpCode(10).Text) = 3 Then
            MsgBox "The Transfer In Division and Union are same as the current Division and Union. Cannot transfer this employee"
            comUserText2.SetFocus
            Exit Function
        End If
    End If
End If 'WFC end

If Len(txtEmpID) = 0 Then
    MsgBox "Employee ID is a required field"
    txtEmpID.SetFocus
    Exit Function
Else
    If Not IsNumeric(txtEmpID) Then
        MsgBox "Invalid Employee ID"
        txtEmpID.SetFocus
        Exit Function
    Else
        If Val(txtEmpID) = 0 Then
            MsgBox "Invalid Employee ID"
            txtEmpID.SetFocus
            Exit Function
        End If
    End If
End If

If lblEmpExist.Visible = True Then
  MsgBox "The NEW Employee number already exits"
  'txtEmpID.SetFocus
  Exit Function
End If

If Len(clpDept) = 0 Then
    MsgBox "Department is a required field"
     clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        MsgBox "Invalid Department"
         clpDept.SetFocus
        Exit Function
    End If
End If

For X% = 2 To 3
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
     clpCode(X%).SetFocus
    Exit Function
End If
Next X%

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

If glbWFC Then 'Ticket #28340 Franks 03/22/2016
    If IsInactivePos(clpJob.Text) Then
        MsgBox "'" & clpJob.Text & "' is Inactive Position Code. Please contact Corporate Total Rewards to review this Position Requirement."
        Exit Function
    End If
    If IsMissingBudPos(clpJob.Text) Then
        MsgBox "Please contact the info:HR corporate administrator to have them create the Budgeted Position Master for '" & clpJob.Text & "' "
        Exit Function
    End If
End If

If Len(medSalary) < 1 Then
    MsgBox "Salary is required"
    medSalary.SetFocus
    Exit Function
End If
If Val(medSalary) <= 0 Then
    MsgBox "Salary is required"
    medSalary.SetFocus
    Exit Function
End If
If comPayPer.ListIndex = -1 Or lblSalCode = "" Then
    MsgBox "Per is required field"
    comPayPer.SetFocus
    Exit Function
End If
If Len(dlpDate(2)) > 0 Then
    If Not IsDate(dlpDate(2)) Then
        MsgBox "Invalid Date of Next Review"
        dlpDate(2).SetFocus
        Exit Function
    End If
Else
    If glbLinamar Then
        MsgBox "Next Review Date is required"
        dlpDate(2).SetFocus
        Exit Function
    End If
End If
If glbLinamar Then
                
    'If Len(txtShift) < 1 Then
    '    MsgBox "Shift is required"
    '    txtShift.SetFocus
    '    Exit Function
    'End If
    
    'Ticket #29414 Franks 11/03/2016 - begin
    If Len(clpCode(13).Text) < 1 Then
        MsgBox lStr("Shift is a required field")
        clpCode(13).SetFocus
        Exit Function
    End If
    If clpCode(13).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
         clpCode(13).SetFocus
        Exit Function
    End If
    'Ticket #29414 Franks 11/03/2016 - end
    
    'If Len(clpCode(4)) < 1 Then
    '    MsgBox "Labour Code is required"
    '    clpCode(4).SetFocus
    '    Exit Function
    'End If
    'Ticket #29946 Franks 03/15/2017
    If Len(txtLabCode.Text) < 1 Then
        MsgBox "Labour Code is required"
        txtLabCode.SetFocus
        Exit Function
    End If
    
End If
If medSIN.Visible Then
    If gSec_Show_SIN_SSN Then
         If Not SIN_chk(medSIN.Text) Then
             MsgBox "Invalid SIN" & IIf(glbLinamar, "", "- if Unassigned set to 999-999-999")
             medSIN.SetFocus
             Exit Function
         End If
    End If
    If CheckSINSSNGen(medSIN, "SIN") Then
        Msg$ = "Duplicate SIN number found. "
        Msg$ = Msg$ & Chr(10) & "Are you sure you wish to accept it?"
        Msg$ = Msg$ & Chr(10) & "Press Yes to accept or No to edit"
        Response% = MsgBox(Msg, MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2, "")    ' Get user response.
        If Response% = IDNO Then
            medSIN.SetFocus
            Exit Function
        End If
    End If
End If
If Not glbLinamar Then
    If glbWFC Then
        If Len(txtPayrollID) = 0 Then
            MsgBox lStr("Payroll ID is a required field.")
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtPayrollID) Then
            MsgBox lStr("Payroll ID must be a number.")
            txtPayrollID.SetFocus
            Exit Function
    End If
    'Ticket #16414
    If xDivCountry = "U.S.A." Then
        If Not (Len(txtPayrollID.Text) = 6) Then
            MsgBox lStr("Payroll ID") & " must be 6 digits for US Divisions"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    
    If Len(dlpDate(3)) > 0 Then
        If Not IsDate(dlpDate(3)) Then
            MsgBox "Invalid Date"
            dlpDate(3).SetFocus
            Exit Function
        End If
    Else
        MsgBox lStr("Original Hire Date is a required field.")
        dlpDate(3).SetFocus
        Exit Function
    End If
    If Len(dlpDate(4)) > 0 Then
        If Not IsDate(dlpDate(4)) Then
            MsgBox "Invalid Date"
            dlpDate(4).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpDate(5)) > 0 Then
        If Not IsDate(dlpDate(5)) Then
            MsgBox "Invalid Date"
            dlpDate(5).SetFocus
            Exit Function
        End If
    End If
    If glbWFC Then
        Call PayIDExist(txtPayrollID, clpCode(0).Text)
        If lblPayIDExist.Visible Then
            Msg$ = "Payroll ID " & txtPayrollID & " already exists in Plant " & clpCode(0).Text
            Msg$ = Msg$ & Chr(10) & " A NEW Payroll ID is required"
            MsgBox Msg$
            txtPayrollID.SetFocus
            Exit Function
        End If
        If Not (ODateOfHire = dlpDate(3).Text) Then
            glbAccessPswd = False
            frmAccessPswd.Show 1
            If glbAccessPswd = False Then   'Access Denied
                MsgBox "Can not change Original Hire Date."
                dlpDate(3).SetFocus
                Exit Function
            End If
        End If
        
        'Ticket #19955 Franks 03/07/2011
        If dlpDate(6).Visible Then
            'If IsDate(xWFC_OldNGSStart) Then
                If Not IsDate(dlpDate(6).Text) Then
                    If Not clpCode(0).Text = "GREN" Then 'Ticket #24268 Franks 10/02/2013
                        If clpCode(12).Text = "COOP" Or clpCode(12).Text = "STUD" Then
                            'Ticket #25352 Franks 04/16/2014 -
                            '"   If Employment Status = COOP or STUD, no NGS Start Date should transfer over
                        Else
                            MsgBox lStr("Other Date 1") & " is a required field."
                            dlpDate(6).SetFocus
                            Exit Function
                        End If
                    End If
                End If
            'End If
        End If
        
        'Ticket #24936 Franks 02/05/2014 - begin
        If Len(elpReptAuthShow(0).Text) = 0 Then
                MsgBox lblReptAuth(0).Caption & " is required"
                elpReptAuthShow(0).SetFocus
                Exit Function
        Else
            If elpReptAuthShow(0).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblReptAuth(0).Caption & " "
                elpReptAuthShow(0).SetFocus
                Exit Function
            End If
        End If
            
        xIsWFCPosChgEmail = False
        If (clpCode(10).Text = "NONE" Or clpCode(10).Text = "EXEC") Then 'Salary employee only 'Ticket #29343 Franks 10/26/2016
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
        
        
        If clpCode(10).Text = "NONE" Then
            If Not IsDate(dlpDate(2).Text) Then
                MsgBox "Next Review is required if Union is 'NONE'"
                dlpDate(2).SetFocus
                Exit Function
            End If
        End If
        If txtFiscalYear.Visible Then
            If Len(txtFiscalYear.Text) = 0 Then
                MsgBox "Fiscal Year is required"
                txtFiscalYear.SetFocus
                Exit Function
            End If
        End If
        If cmbMarketLine.Visible Then
            If Len(cmbMarketLine.Text) = 0 Then
                MsgBox "Market Line is required"
                cmbMarketLine.SetFocus
                Exit Function
            End If
        End If
        'Ticket #24936 Franks 02/05/2014 - end
    End If ' end of WFC
    
End If

If glbLinamar Then 'Ticket #29414 Franks 11/03/2016
    If Len(elpReptAuthShow(0).Text) = 0 Then
        MsgBox lStr("Rept. Authority 1") & " is required."
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
    If Len(elpReptAuthShow(0)) > 0 Then
        If elpReptAuthShow(0).Caption = "Unassigned" Then
            MsgBox lStr("Rept. Authority 1") & " not valid. Check Employee # and re-enter!"
            elpReptAuthShow(0).SetFocus
            Exit Function
        End If
    End If
End If

ChkInput = True

End Function

Private Sub clpBGroup_LostFocus()
If glbWFC Then 'Ticket #23247 Franks 09/13/2013
    Call getValsFromBenGrpMatrix(clpBGroup.Text, clpDIV.Text)
End If
End Sub

Private Sub clpCode_Change(Index As Integer)
    If Index = 6 And glbWFC And cmbMarketLine.Visible = True Then
        Call Set_MarketLine_List
    End If
    
    'Ticket #24451 Franks 10/17/2013
    If glbWFC And Index = 10 Then 'Union
        If Len(clpCode(10).Text) > 0 Then
            'If Not clpCode(10).Caption = "Unassigned" Then
                If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/17/2014
                    'dont change
                Else
                    Call WFCDefaultHours(clpCode(10).Text)
                End If
            'End If
        End If
    End If
End Sub

Private Sub WFCDefaultHours(xUnion) 'Ticket #24451 Franks 10/17/2013
    If Len(xUnion) > 0 Then
        If (xUnion = "NONE" Or xUnion = "EXEC" Or xUnion = "-NON" Or xUnion = "-EXE") Then  'salaried
            medHours(0).Text = 8
            medHours(1).Text = 40
            medHours(2).Text = 86.67
        Else 'hourly
            medHours(0).Text = 8
            medHours(1).Text = 40
            medHours(2).Text = 40
        End If
    End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If glbWFC Then
        If Index = 10 Then 'Union
            Call WFC_UnionScreen(clpCode(10).Text)
            Call WFC_Band 'Ticket #21677 Franks 03/07/2012
            'Call getPayGroup(clpDIV.Text, clpCode(10).Text)
        End If
        If Index = 10 Or Index = 12 Then 'Union & Status
            Call DispNGSBenGroups 'Ticket #23247 Franks 09/13/2013
        End If
    End If
End Sub

Private Sub clpDept_LostFocus()
Call Dept_GL
End Sub

Private Sub clpDiv_LostFocus()
    If Len(clpDIV) > 0 And clpDIV.Caption <> "Unassigned" Then
        If InStr(lblEmpNo, "-") > 0 Then
            lblEmpNo.Caption = clpDIV & Mid(lblEmpNo, InStr(lblEmpNo, "-"))
        End If
        If glbWFC Then 'Ticket #18654
            If GetCountryFromDiv(clpDIV.Text) = "CANADA" Then
                lblBen.FontBold = True
            Else
                lblBen.FontBold = False
            End If
        End If
    End If
    If Len(clpDIV) > 0 And clpDIV.Caption <> "Unassigned" And InStr(lblEmpNo, "-") > 3 Then
        If InStr(lblEmpNo, "-") > 0 Then
            lblEmpNo.Caption = clpDIV & Mid(lblEmpNo, InStr(lblEmpNo, "-"))
        End If
        If glbLinamar Then
        Call set_CountrySIN(clpDIV)
        End If
    End If
End Sub

Private Sub clpJob_Change()
Call WFC_Band 'Ticket #21677 Franks 03/07/2012
Call Set_MarketLine_List 'Ticket #24620 Franks 12/03/2013
End Sub

Private Sub clpJob_LostFocus()
Call Job_Desc
Call WFC_Band 'Ticket #21677 Franks 03/07/2012

If glbWFC Then 'Ticket #29343 Franks 10/26/2016
    Dim xStr
    If Len(clpJob.Text) > 0 Then
        xStr = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
        If Len(elpReptAuthShow(0).Text) = 0 And Len(clpJob.Text) > 0 Then
            elpReptAuthShow(0).Text = xStr 'GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
        End If
        If Len(xStr) = 0 Then
            If (clpCode(10).Text = "NONE" Or clpCode(10).Text = "EXEC") Then 'Salary employee only
                lblWFCNote.Visible = True
            End If
        Else
            lblWFCNote.Visible = False
        End If
    End If
End If

End Sub

Private Sub WFC_Band()
    If glbWFC Then
        fglbBAND = getPosBand(clpJob.Text)
        If Len(fglbBAND) > 0 Then
            lblBand.Caption = "Band: " & fglbBAND
        Else
            lblBand.Caption = ""
        End If
        lblBand.Top = lblFiscalYear.Top
        lblBand.Visible = lblMarketLine.Visible
    End If
End Sub
Public Sub cmdClose_Click()

On Error GoTo err_Unload

glbTERM_ID = 0
glbTran_ID = 0
glbTran_Seq = 0
glbOnTop = ""

Unload Me

Exit Sub

err_Unload:
Unload Me
Resume Next
Unload Me

End Sub


Private Sub cmdEditPayID_Click()
    glbAccessPswd = False
    frmAccessPswd.Show 1
    If glbAccessPswd = False Then   'Access Denied
        Exit Sub
    End If
    xPayIDEnable = True
    txtPayrollID.Enabled = True
    txtPayrollID.SetFocus
End Sub

Private Sub cmdFrankTest_Click()
    If IsWFCReptPosAuth(clpJob.Text) Then
        If IsDate(dlpDate(0).Text) Then
            glbWFCNewPosJob = clpJob.Text
            glbWFC_IncePlanID = txtEmpID.Text 'glbLEE_ID
            glbWFC_IPPopFormName = "WFCEmpListWithRepTranIn"
            glbWFC_CancelTransaction = False
            frmCheckListView.lblStDate = dlpDate(0).Text  ' dlpTermDate.Text
            frmCheckListView.Show 1
        End If
    End If
End Sub

Public Sub cmdOK_Click()
Dim Msg$, Title$, DgDef As Variant, Response%
Dim EID&, SEQID&, TermDate$, X%
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
If glbTran_ID = 0 Then Exit Sub
If glbSamuel Then
    If Not ChkInSamuel() Then Exit Sub
Else
    If Not ChkInput() Then Exit Sub
End If

If glbWFC Then 'Ticket #25248 Franks 03/24/2014
    If glbDivTranInPlant = "Y" Then
        Msg$ = Msg$ & Chr(10) & "The default values were displayed for this employee. "
        Msg$ = Msg$ & Chr(10) & "Click 'No' to return to the Transfer In screen to make the changes."
        Msg$ = Msg$ & Chr(10) & "Click 'Yes' to continue with the transfer."
    Else
        Msg$ = Msg$ & Chr(10) & "Are you sure you want to transfer this employee ?"
    End If
Else
    Msg$ = Msg$ & Chr(10) & "Are you sure you want to transfer this employee ?"
End If
'Ticket #24184 Franks 11/14/2013, Jerry asked to remove the followings
'Msg$ = Msg$ & Chr(10) & "this employee ?"
'Msg$ = Msg$ & Chr(10) & "Make sure no other info:HR Window "
'Msg$ = Msg$ & Chr(10) & "is open with this employee information showing"

Title$ = "Transfer In Employee"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodType = 1
If glbSamuel Then
    EID& = txtEmpID.Text 'ShowEmpnbr(glbTran_ID)
Else
    EID& = CLng(getEmpnbr(lblEmpNo))
End If
SEQID& = CLng(fglbTERM_Seq)
If medHours(0) = "" Then medHours(0) = 0
If medHours(1) = "" Then medHours(1) = 0
If medHours(2) = "" Then medHours(2) = 0

If IsWFCReptPosAuth(clpJob.Text) Then 'Ticket #29507 Franks 12/01/2016
    If IsDate(dlpDate(0).Text) Then
        glbWFCNewPosJob = clpJob.Text
        glbWFC_IncePlanID = txtEmpID.Text 'glbLEE_ID
        glbWFC_IPPopFormName = "WFCEmpListWithRepTranIn"
        glbWFC_CancelTransaction = False
        frmCheckListView.lblStDate = dlpDate(0).Text  ' dlpTermDate.Text
        frmCheckListView.Show 1
        If glbWFC_CancelTransaction Then
            Exit Sub
        End If
    End If
End If
    
If glbWFC And glbCandidate > 0 Then 'Ticket #24184 Franks 10/28/2013
    ''Call WFCHRSoftTransferOut(clpDIV.Text, clpCode(10).Text, dlpDate(1).Text)
    Call WFCHRSoftTermination(EID&, DateAdd("D", -1, CVDate(dlpDate(1).Text)), clpDIV.Text, clpCode(10).Text, False)
    
    glbTran_Seq = glbTERM_Seq
    Call EERetrieveHRSoftAfterTransferOut(glbTran_Seq)
    
    'Exit Sub 'will remove later
    'HRSoft Transfer - not delete the employee info from HREMP, just moved the data to Term tables,
    'so do not need to do the following steps
    'If Not modReinMove(EID&, SEQID&, TermDate) Then Exit Sub
Else
    If Not modReinMove(EID&, SEQID&, TermDate) Then Exit Sub
End If

MDIMain.panHelp(0).FloodPercent = 100
Call Upd_Related_EMP(EID&)
Call Upd_Related_Job(EID&)
Call Upd_Related_Salary(EID&)
If glbWFC Then 'Ticket #24317 Franks 09/16/2013
    If Not clpDept.Text = OldDept Then
        If Not EmpHisCalc(2, EID&, clpDept, "", "", "", "", "", "", dlpDate(0), , , , , , , OldDept) Then MsgBox "EMPHIS Error "
    End If
    If Not clpDIV.Text = OldDiv Then
        If Not EmpHisCalc(2, EID&, "", clpDIV, "", "", "", "", "", dlpDate(0), , , , , , , OldDiv) Then MsgBox "EMPHIS Error "
    End If
    If Not clpCode(10).Text = OldUnion Then
        If Not EmpHisCalc(2, EID&, "", "", "", "", clpCode(10).Text, "", "", dlpDate(0), , , , , , , OldUnion) Then MsgBox "EMPHIS Error "
    End If
    'Ticket #25088 Franks 02/18/2014 - begin
    If Not clpCode(0).Text = OldSection Then
        If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", dlpDate(0).Text, "SECTION", clpCode(0).Text, , , , , OldSection) Then MsgBox "EMPHIS Error "
    End If
    If Not clpCode(6).Text = OldRegion Then
        If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", dlpDate(0).Text, "REGION", clpCode(6).Text, , , , , OldRegion) Then MsgBox "EMPHIS Error "
    End If
    If Not clpCode(5).Text = OldAdminBy Then
        If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", dlpDate(0).Text, "ADMINBY", clpCode(5).Text, , , , , OldAdminBy) Then MsgBox "EMPHIS Error "
    End If
    If Not clpCode(7).Text = OldLoc Then
        If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", dlpDate(0).Text, "LOC", clpCode(7).Text, , , , , OldLoc) Then MsgBox "EMPHIS Error "
    End If
    'Ticket #25088 Franks 02/18/2014 - end
Else
    If Not EmpHisCalc(2, EID&, clpDept, "", "", "", "", "", "", dlpDate(0)) Then MsgBox "EMPHIS Error "
    If Not EmpHisCalc(2, EID&, "", clpDIV, "", "", "", "", "", dlpDate(0)) Then MsgBox "EMPHIS Error "

    If glbLinamar Then
        If clpCode(2) <> "" Then If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", Date, "REGION", getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)) Then MsgBox "EMPHIS Error "
    Else
        If clpCode(2) <> "" Then If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", Date, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
    End If
    If clpCode(3) <> "" Then If Not EmpHisCalc(2, EID&, "", "", "", "", "", "", "", Date, "SECTION", clpCode(3)) Then MsgBox "EMPHIS Error "

End If

X% = modReinAudit(EID&)
X% = RehHREMPAudit(EID&)
X% = RehJOBAudit(EID&)
X% = RehSALARYAudit(EID&)
X% = RehBENEFITSAudit(EID&)
Call updFollow(EID&, "R")
''''Added by Bryan 04/JAN/06 Ticket#10066
'''If EID <> glbTran_ID Then
'''    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR=" & EID
'''End If
'''gdbAdoIhr001.Execute "UPDATE HR_PHOTO SET PT_EMPNBR=" & EID& & " WHERE PT_EMPNBR=" & glbTran_ID
'Ticket #24552 Franks 11/01/2013
If EID <> xTL_OLDEMPNBR Then
    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR=" & EID
    gdbAdoIhr001.Execute "UPDATE HR_PHOTO SET PT_EMPNBR=" & EID& & " WHERE PT_EMPNBR=" & xTL_OLDEMPNBR
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") + 1
rsT_PARCO.Update
rsT_PARCO.Close
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


If glbWFC Then

    'Ticket #19266 Franks 11/29/10
    Call WFC_NGS_Trans(EID&, "Transfer In")

    'Ticket #21677 Franks 03/15/2012
    'Ticket #23247 Franks 09/16/2013 - moved this function before locBeneGroupUpdate
    Call locBeneGroupUpdate(EID&)
    
    Call updBenefitForSalDEPN(EID&, , dlpDate(1).Text) 'Ticket #23247 Franks 09/16/2013
    
    'Frank 03/06/2009 Ticket #16234
    Call Employee_Master_Integration(EID&)
    'Frank 09/16/2009 Ticket #16395
    'Call WFCPensionMaster(EID&, "Y")
    toSOURCE = "IHR Transfer In" 'Ticket #19954
    'Ticket #21786 Franks - use Transfer Date instead of Division Start Daet (dlpDate(0) => dlpDate(1))
    If Len(clpCode(10).Text) > 0 Then 'Ticket #21677 Franks 03/14/2012 - Union Transfer
        Call WFCPensionMasUpt(EID&, "Transfer In", dlpDate(1).Text, Left(OldDiv, 4), Year(CVDate(dlpDate(0).Text)), , OldUnion)
    Else
        Call WFCPensionMasUpt(EID&, "Transfer In", dlpDate(1).Text, Left(OldDiv, 4), Year(CVDate(dlpDate(0).Text)))
    End If
    
    ''Ticket #19266 Franks 11/29/10
    'Call WFC_NGS_Trans(EID&, "Transfer In")
    
    'Ticket #23247 Franks 09/13/2013 - benefit group change
    
    If glbCandidate > 0 Then 'Ticket #24184 Franks 11/13/2013
        Call WFCHRSoftProcUpt("frmETRANIN", EID&)
    End If
    
    Call mod_Upd_Pos_Budget_WFC(clpJob.Text, "") 'Ticket #25911 Franks 12/17/2014
    
    If xIsWFCPosChgEmail Then  'Ticket #29343 Franks 10/26/2016
        If gsEMAIL_ONPOSITION Then
            If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
                Call WFCPubPosChangedcmdEmail(EID&, xWFCPosChgEmailBody, "info:HR Position Reporting Authority Change Notice")
            End If
        End If
    End If
        
End If

If glbLinamar Then 'Ticket #29759 Franks 02/22/2017
    If xPayIDEnable Then
        Call getNextLinPayrollID("Y")
    End If
End If
MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT

lblEENum = 0
lblEmpNo = 0
 clpDIV = ""
txtEmpID = ""
 clpDept = ""
For X% = 0 To 2
    dlpDate(X%) = ""
     clpCode(X% + 1) = ""
     clpHOME(X% + 1) = ""
    medHours(X%) = ""
Next
txtShift = ""
If glbLinamar Then
    clpCode(13).Text = 0
End If
 clpCode(4) = ""
 clpHOME(4) = ""
 clpGLNum = ""
 clpJob = ""
medSalary = ""
medSIN = ""
lblCountry = ""
lblPROV = ""
Call UPDMOD
lblEEName = "Employee was Transferred"
fglbTERM_Seq = 0
glbTERM_ID = 0
glbTran_ID = 0
glbTran_Seq = 0
Unload Me
End Sub

Private Function EERetrievWFCTran() 'Ticket #24184 Franks 10/28/2013
Dim SQLQ As String
EERetrievWFCTran = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

'SQLQ = SQLQ & "SELECT LN_TRALOG.*, Term_HREMP.ED_ORG,Term_HREMP.ED_EMPTYPE,Term_HREMP.ED_WORKCOUNTRY,Term_HREMP.ED_BENEFIT_GROUP "
'SQLQ = SQLQ & ",Term_HREMP.ED_SURNAME,Term_HREMP.ED_FNAME,Term_HREMP.ED_DEPTNO"
'SQLQ = SQLQ & " FROM LN_TRALOG INNER JOIN Term_HREMP ON LN_TRALOG.TL_TERM_SEQ = Term_HREMP.TERM_SEQ "
'SQLQ = SQLQ & "WHERE TL_TERM_SEQ = " & glbTran_Seq
SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbTran_ID
If rsORG.State <> 0 Then rsORG.Close
rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic

If glbWFC Then 'Create by Bryan Ticket#12510 Feb 06/07
    Dim rs As New ADODB.Recordset
    fglbBAND = ""
    'SQLQ = "SELECT HRJOB.JB_BAND FROM LN_TRALOG INNER JOIN Term_JOB_HISTORY ON "
    'SQLQ = SQLQ & "LN_TRALOG.TL_TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ INNER JOIN HRJOB ON Term_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    'SQLQ = SQLQ & "Where (Term_JOB_HISTORY.JH_CURRENT <> 0) and  TL_TERM_SEQ = " & glbTran_Seq
    SQLQ = "SELECT HRJOB.JB_BAND FROM HREMP INNER JOIN HR_JOB_HISTORY ON "
    SQLQ = SQLQ & "HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "Where (HR_JOB_HISTORY.JH_CURRENT <> 0) AND ED_EMPNBR = " & glbTran_ID
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If IsNull(rs("JB_BAND")) = False Then
            fglbBAND = rs("JB_BAND")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If


EERetrievWFCTran = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TRANSFER", "HREMP", "SELECT")

Resume Next

Exit Function


End Function
Private Function EERetrieveHRSoftAfterTransferOut(xTran_Seq)
Dim SQLQ As String
'EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

SQLQ = SQLQ & "SELECT LN_TRALOG.*, Term_HREMP.ED_ORG,Term_HREMP.ED_EMPTYPE,Term_HREMP.ED_WORKCOUNTRY,Term_HREMP.ED_BENEFIT_GROUP "
SQLQ = SQLQ & ",Term_HREMP.ED_SURNAME,Term_HREMP.ED_FNAME,Term_HREMP.ED_DEPTNO"
SQLQ = SQLQ & " FROM LN_TRALOG INNER JOIN Term_HREMP ON LN_TRALOG.TL_TERM_SEQ = Term_HREMP.TERM_SEQ "
SQLQ = SQLQ & "WHERE TL_TERM_SEQ = " & xTran_Seq
If rsORG.State <> 0 Then rsORG.Close

rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
If Not rsORG.EOF Then xTL_OLDEMPNBR = rsORG("TL_OLDEMPNBR") Else xTL_OLDEMPNBR = 0

If glbWFC Then 'Create by Bryan Ticket#12510 Feb 06/07
    Dim rs As New ADODB.Recordset
    fglbBAND = ""
    SQLQ = "SELECT HRJOB.JB_BAND FROM LN_TRALOG INNER JOIN Term_JOB_HISTORY ON "
    SQLQ = SQLQ & "LN_TRALOG.TL_TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ INNER JOIN HRJOB ON Term_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "Where (Term_JOB_HISTORY.JH_CURRENT <> 0) and  TL_TERM_SEQ = " & xTran_Seq
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If IsNull(rs("JB_BAND")) = False Then
            fglbBAND = rs("JB_BAND")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If


'EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REHIRE", "Term_HRTRMEMP", "SELECT")

Resume Next


End Function

Private Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

SQLQ = SQLQ & "SELECT LN_TRALOG.*, Term_HREMP.ED_ORG,Term_HREMP.ED_EMPTYPE,Term_HREMP.ED_WORKCOUNTRY,Term_HREMP.ED_BENEFIT_GROUP "
SQLQ = SQLQ & ",Term_HREMP.ED_SURNAME,Term_HREMP.ED_FNAME,Term_HREMP.ED_DEPTNO"
SQLQ = SQLQ & ",Term_HREMP.ED_SECTION,Term_HREMP.ED_REGION,Term_HREMP.ED_ADMINBY,Term_HREMP.ED_LOC" 'Ticket #25088 Franks 02/18/2014
SQLQ = SQLQ & ",Term_HREMP.ED_SIN" 'Ticket #29759 Franks 02/22/2017
SQLQ = SQLQ & " FROM LN_TRALOG INNER JOIN Term_HREMP ON LN_TRALOG.TL_TERM_SEQ = Term_HREMP.TERM_SEQ "
SQLQ = SQLQ & "WHERE TL_TERM_SEQ = " & glbTran_Seq
If rsORG.State <> 0 Then rsORG.Close

rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic

If glbWFC Then 'Create by Bryan Ticket#12510 Feb 06/07
    Dim rs As New ADODB.Recordset
    fglbBAND = ""
    SQLQ = "SELECT HRJOB.JB_BAND FROM LN_TRALOG INNER JOIN Term_JOB_HISTORY ON "
    SQLQ = SQLQ & "LN_TRALOG.TL_TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ INNER JOIN HRJOB ON Term_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "Where (Term_JOB_HISTORY.JH_CURRENT <> 0) and  TL_TERM_SEQ = " & glbTran_Seq
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If IsNull(rs("JB_BAND")) = False Then
            fglbBAND = rs("JB_BAND")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If

If glbSamuel Then 'Ticket #21791 Franks 04/02/2012
    If Not rsORG.EOF Then
        If Not IsNull(rsORG("TL_NEWPLANT")) Then
            clpCode(5).Text = rsORG("TL_NEWPLANT")
        End If
    End If
End If

xLocSIN = ""
'If glbLinamar Then 'Ticket #29759 Franks 02/22/2017
    If Not rsORG.EOF Then
        If Not IsNull(rsORG("ED_SIN")) Then
            xLocSIN = rsORG("ED_SIN")
        End If
    End If
'End If

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REHIRE", "Term_HRTRMEMP", "SELECT")

Resume Next

Exit Function

End Function
Private Sub PayIDExist(xPayID, xPlantCode)
Dim SQLQ
Dim rsEmp As New ADODB.Recordset
If xPayID = "" Then xPayID = 0
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP"
SQLQ = SQLQ & " WHERE HREMP.ED_PAYROLL_ID = '" & xPayID & "' "
SQLQ = SQLQ & " AND ED_SECTION = '" & xPlantCode & "' "
If glbCandidate > 0 Then 'Ticket #24184 Franks 10/28/2013
    SQLQ = SQLQ & " AND NOT (ED_EMPNBR = " & glbTran_ID & ") "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEmp.EOF And rsEmp.BOF Then
    lblPayIDExist.Caption = ""
    lblPayIDExist.Visible = False
    'Ticket #20708 Franks 07/27/2011, only check active employees
    ''rsEmp.Close
    ''SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP"
    ''SQLQ = SQLQ & " WHERE Term_HREMP.ED_PAYROLL_ID = '" & xPayID & "' "
    ''SQLQ = SQLQ & " AND ED_SECTION = '" & xPlantCode & "' "
    ''rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    ''If Not (rsEmp.EOF And rsEmp.BOF) Then
    ''    lblPayIDExist.Caption = " Payroll ID " & xPayID & " already exists - A NEW Payroll ID is required"
    ''    lblPayIDExist.Visible = True
    ''End If
Else
    lblPayIDExist.Caption = " Payroll ID " & xPayID & " already exists - A NEW Payroll ID is required"
    lblPayIDExist.Visible = True
End If
rsEmp.Close

End Sub
Private Sub EmpNoExist(xEMP)
Dim SQLQ
Dim rsEmp As New ADODB.Recordset

If glbWFC And glbCandidate > 0 Then 'Ticket #24184 Franks 10/28/2013
    Exit Sub
End If

If xEMP = "" Then xEMP = 0
SQLQ = "SELECT ED_EMPNBR FROM HREMP"
SQLQ = SQLQ & " where HREMP.ED_EMPNBR = " & xEMP
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEmp.EOF And rsEmp.BOF Then
    lblEmpExist.Caption = ""
    lblEmpExist.Visible = False
    If glbLinamar Then
        If EmpNoInTerm(xEMP) Then
            lblEmpExist.Caption = " Employee # " & ShowEmpnbr(xEMP) & " already active in terminated list - A NEW Number is required"
            lblEmpExist.Visible = True
        End If
    End If
Else
    lblEmpExist.Caption = " Employee # " & ShowEmpnbr(xEMP) & " already active - A NEW Number is required"
    lblEmpExist.Visible = True
End If
End Sub

Private Sub comPayPer_GotFocus()
    Call SetPanHelp(ActiveControl)
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

Private Sub comUserText2_Click()
If glbWFC Then 'Ticket #22448 Franks
    If comUserText2.ListIndex <> -1 Then
        txtUserText2.Text = getUserText2(comUserText2.Text)
    End If
End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMETRANIN"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMETRANIN"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTERM As New ADODB.Recordset
Dim X%, SQLQ

glbOnTop = "FRMETRANIN"

Me.WindowState = vbMaximized
If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    Me.Left = 0
End If
Call setCaption(lbltitle(1))
Call setCaption(lbltitle(21))
Call setCaption(lbltitle(6))
Call setCaption(lbltitle(7))
Call setCaption(lbltitle(10))
Call setCaption(lbltitle(13))
Call setCaption(lbltitle(23))
Call setCaption(lbltitle(24))
Call setCaption(lbltitle(25))
lbltitle(28).Caption = lStr("Pay Period") 'Ticket #21988 Franks 05/02/2012

comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "
'woodbridge get's Daily salary - Ticket #14645
If glbCompSerial = "S/N - 2282W" Or glbLinamar Then
    comPayPer.AddItem "Daily "
End If
Call DecSetup

If glbWFC Then 'Ticket #25911 Franks 12/17/2014
    clpJob.TextBoxWidth = 1215
    clpJob.TransDiv = glbWFCUserSecList
End If

'Ticket #24184 Franks 10/23/2013 -  HSSoft Transfer In
If glbWFC And glbCandidate > 0 Then
    Call WFCHRSoftDispValues
    Exit Sub
End If

If glbWFC And glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/18/2014
    Call WFCDivTranInSamePlant
Else
    frmTranEMPL.Show 1
End If
If glbTran_ID = 0 Then
    Unload Me
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If Len(glbTran_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Transfer In - " & Left$(glbTran_SName, 5)
    Me.lblEEName = RTrim$(glbTran_SName) & ", " & RTrim$(glbTran_Fname)
End If

lblEENum.Caption = ShowEmpnbr(glbTran_ID)

If EERetrieve() = False Then Exit Sub

If Not rsORG.EOF Then
    clpDIV = rsORG("TL_NEWDIV")
    xTL_OLDEMPNBR = rsORG("TL_OLDEMPNBR") 'Ticket #24552 Franks 11/01/2013
End If
If glbLinamar Then
    fglbTERM_Seq = rsORG("TL_TERM_SEQ")
    Call LinamarScreenSetup 'Ticket #29414 Franks 11/03/2016
ElseIf glbSamuel Then
    Call SamuelScreenSetup
Else 'For WFC
    Call WFCScreenSetup 'Ticket #22448
End If
lblEmpNo = ShowEmpnbr(glbTran_ID)

'If Not rsORG.EOF Then
    If glbWFC Then
        If IsDate(rsORG("TL_NEWDIVEDATE")) Then
            dlpDate(0) = DateAdd("D", 1, rsORG("TL_NEWDIVEDATE"))
        End If
        'Ticket #21677 Franks 03/14/2012
        If Not IsNull(rsORG("TL_NEW_ORG")) Then
            clpCode(10).Text = rsORG("TL_NEW_ORG")
            Call WFC_UnionScreen(rsORG("TL_NEW_ORG"))
            Call getPayGroup(rsORG("TL_NEWDIV"), rsORG("TL_NEW_ORG"))
        End If

    Else
        dlpDate(0) = rsORG("TL_NEWDIVEDATE")
    End If
    fglbTERM_Seq = rsORG("TL_TERM_SEQ")
'End If
dlpDate(1) = dlpDate(0)

If glbWFC Then Call WFC_DispNGSStartDate 'Ticket #24652 Franks 12/02/2013

clpGLNum.TextBoxWidth = 1200
Call Set_PositionSalary
If glbWFC Then
    ODateOfHire = dlpDate(3).Text
End If
xUpdateable = True
If glbLinamar Then
Call set_CountrySIN(clpDIV)
End If
MDIMain.panHelp(1).Caption = " "

Call INI_Controls(Me)
End Sub
Private Sub WFC_PosSalScreen() 'Ticket #24936 Franks 02/05/2014
Dim xVal
    xVal = 360
    If Not glbWFC Then Exit Sub
    
    lblReptAuth(0).Top = lbltitle(20).Top
    lblReptAuth(0).Left = lbltitle(20).Left
    elpReptAuthShow(0).Top = medHours(0).Top
    elpReptAuthShow(0).Left = clpJob.Left
    
    lblWFCNote.Top = elpReptAuthShow(0).Top 'Ticket #29343 Franks 10/26/2016
    lblWFCNote.Left = elpReptAuthShow(0).Left + 3500
    'lblWFCNote.Visible = True '???
    
    lblReptAuth(0).FontBold = True
    lblReptAuth(0).Visible = True
    elpReptAuthShow(0).Visible = True
    elpReptAuthShow(0).TabIndex = medHours(0).TabIndex
    
    lbltitle(20).Top = lbltitle(20).Top + xVal
    lbltitle(17).Top = lbltitle(17).Top + xVal
    lbltitle(15).Top = lbltitle(15).Top + xVal
    lbltitle(14).Top = lbltitle(14).Top + xVal
    lbltitle(22).Top = lbltitle(22).Top + xVal
    lbltitle(5).Top = lbltitle(5).Top + xVal
    lblEEStatus.Top = lblEEStatus.Top + xVal
    
    medHours(0).Top = medHours(0).Top + xVal
    medHours(1).Top = medHours(1).Top + xVal
    medHours(2).Top = medHours(2).Top + xVal
    medSalary.Top = medSalary.Top + xVal
    comPayPer.Top = comPayPer.Top + xVal
    lblSalCode.Top = lblSalCode.Top + xVal
    dlpDate(2).Top = dlpDate(2).Top + xVal
    clpCode(8).Top = clpCode(8).Top + xVal
'    lblEEStatus.Top = lblEEStatus.Top + xVal
'    lblEEStatus.Top = lblEEStatus.Top + xVal
'    lblEEStatus.Top = lblEEStatus.Top + xVal
    
    
End Sub
Private Sub WFC_UnionScreen(xORG)
Dim xVal
    xVal = 360
        If Not glbWFC Then Exit Sub
        If xORG = "NONE" Or xORG = "EXEC" Then
            txtFiscalYear.Visible = True
            txtFiscalYear.Top = 5925 + xVal
            cmbMarketLine.Visible = True
            cmbMarketLine.Top = 6270 + xVal
            lblFiscalYear.Visible = True
            lblFiscalYear.Top = 5925 + xVal
            lblMarketLine.Visible = True
            lblMarketLine.Top = 6270 + xVal
            lblMLine.Visible = True
            lblMLine.Top = 6270 + xVal
            lbltitle(28).Top = 6645 + xVal '#15818
            clpCode(9).Top = 6645 + xVal '#15818
            lblVadim2.Top = clpCode(9).Top + 360 '+ xVal
            clpVadim2.Top = clpCode(9).Top + 360 '+ xVal
            ''Ticket #23247 Franks 09/13/2013
            'lblVadim11.Left = lblVadim2.Left
            'lblVadim11.Top = lblVadim2.Top + 360
            'clpVadim1.Left = clpVadim2.Left
            'clpVadim1.Top = clpVadim2.Top + 360
        Else
            'Hiding some NONE and EXEC fields
            txtFiscalYear.Visible = False
            cmbMarketLine.Visible = False
            lblFiscalYear.Visible = False
            lblMarketLine.Visible = False
            lblMLine.Visible = False
            
            lbltitle(28).Top = 5940 + xVal '#15818
            clpCode(9).Top = 5940 + xVal '#15818
            lblVadim2.Top = clpCode(9).Top + 360
            clpVadim2.Top = clpCode(9).Top + 360
            ''Ticket #23247 Franks 09/13/2013
            'lblVadim11.Left = lblVadim2.Left
            'lblVadim11.Top = lblVadim2.Top + 360
            'clpVadim1.Left = clpVadim2.Left
            'clpVadim1.Top = clpVadim2.Top + 360
        End If
        
        'Ticket #30446 Franks 08/09/2017
        lbltitle(31).Top = lblVadim2.Top + 360
        clpCode(14).Top = clpVadim2.Top + 360
End Sub


Private Function modReinMove(EID&, EESEQ&, TermDate$)
Dim X%, DtTm   As Variant, TRDesc$

Screen.MousePointer = HOURGLASS
modReinMove = False
DtTm = Now

MDIMain.panHelp(0).FloodPercent = 5

If glbWFC Then
    X% = REIN_BASIC(EID&, EESEQ&, TermDate$, txtPayrollID)
Else
    X% = REIN_BASIC(EID&, EESEQ&, TermDate$)
End If

If Not X% Then
    Exit Function
End If

'Ticket #19488 Frank 11/29/10
X% = REIN_HREMP_OTHER(EID&, EESEQ&)

X% = REIN_DEPEND(EID&, EESEQ&)

X% = REIN_COBRA(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 10

X% = REIN_ATTENDANCE(EID&, EESEQ&)

X% = REIN_BENEFITS(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 20

X% = REIN_HealthCost(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 25

X% = REIN_HealthSafety(EID&, EESEQ&)

X% = REIN_JOB(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 40

X% = REIN_PERFORM(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 60

X% = REIN_SALARY(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 75

X% = REIN_EDUCSEM(EID&, EESEQ&)
X% = REIN_COMMENTS(EID&, EESEQ&)
X% = REIN_EARN(EID&, EESEQ&)
X% = REIN_EDU(EID&, EESEQ&)
X% = REIN_EMPSKL(EID&, EESEQ&)
X% = REIN_TRADE(EID&, EESEQ&)
X% = REIN_DOLENT(EID&, EESEQ&)

'Ticket #28789 - Actual Amounts Details
X% = REIN_DOLENT_ACTDTL(EID&, EESEQ&)

X% = REIN_ENTHRS(EID&, EESEQ&)
X% = REIN_COUNSEL(EID&, EESEQ&)
If glbLinamar Then
    X% = REIN_LN_EMPSKL(EID&, EESEQ&)
End If

'Ticket #17894 01/20/10
'add the missed tables - begin
MDIMain.panHelp(0).FloodPercent = 80
X% = REIN_OHS_CONTACT(EID&, EESEQ&)
X% = REIN_OHS_CORRECTIVE(EID&, EESEQ&)
X% = REIN_OHS_ROOT_CAUSES(EID&, EESEQ&)
X% = REIN_OHS_CLAIM_MEDICAL(EID&, EESEQ&)
X% = REIN_USERDEFINED(EID&, EESEQ&)
X% = REIN_EDU(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 85
X% = REIN_SUCCESSION(EID&, EESEQ&)
X% = REIN_LANGUAGE(EID&, EESEQ&)
X% = REIN_HREMPHIS(EID&, EESEQ&)
X% = REIN_Profit_Sharing(EID&, EESEQ&)
X% = REIN_EMP_FLAGS(EID&, EESEQ&)     'Hemu - Ticket #24065
X% = REIN_HREEO(EID&, EESEQ&) 'Ticket #25669 Franks 06/24/2014
Call DelTermEEO(EESEQ&) 'Ticket #25669 Franks 06/24/2014

MDIMain.panHelp(0).FloodPercent = 90
If gsAttachment_DB Then
    X% = REIN_HRDOC_EMP(EID&, EESEQ&)
    X% = REIN_HRDOC_JOB_HISTORY(EID&, EESEQ&)
    X% = REIN_HRDOC_COMMENTS(EID&, EESEQ&)
    X% = REIN_HRDOC_HEALTH_SAFETY(EID&, EESEQ&)
    X% = REIN_HRDOC_HEALTH_SAFETY_2(EID&, EESEQ&)
    X% = REIN_HRDOC_COUNSEL(EID&, EESEQ&)
    X% = REIN_HRDOC_PERFORM_HISTORY(EID&, EESEQ&)
    X% = REIN_HRDOC_EDSEM(EID&, EESEQ&)
    X% = REIN_HRDOC_EDSEM_RETEST(EID&, EESEQ&)
    X% = REIN_HRDOC_HREDU(EID&, EESEQ&)
    X% = REIN_HRDOC_DOLENT(EID&, EESEQ&)
End If
'add the missed tables - end

modReinMove = True

Screen.MousePointer = DEFAULT

Exit Function

modReinMoveErr_Msg:
Screen.MousePointer = DEFAULT
MsgBox "Problem Creating Audit record - Termination Aborted"

End Function

Private Function RehBENEFITSAudit(EEID&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
RehBENEFITSAudit = False

On Error GoTo RehBENEFITSAudit_Err

SQLQ = "INSERT INTO HRAUDIT (AU_COMPNO, AU_EMPNBR, AU_BCODE, "
SQLQ = SQLQ & "AU_COVER, AU_BAMT, AU_EDATE, AU_PCE, AU_PCC, "
SQLQ = SQLQ & "AU_TCOST, AU_UNITCOST, AU_PREMIUM, AU_PER, "
SQLQ = SQLQ & "AU_PPAMT, AU_MTHCCOST, AU_MTHECOST,AU_TAXBEN, "
If glbWFC Then SQLQ = SQLQ & "AU_VSTEP, " 'Ticket #25275 Franks 04/02/2014
SQLQ = SQLQ & "AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE,AU_UPLOAD ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)

SQLQ = SQLQ & "SELECT BF_COMPNO, " & EEID & " , BF_BCODE, "
SQLQ = SQLQ & "BF_COVER, BF_AMT, BF_EDATE, BF_PCE, BF_PCC, "
SQLQ = SQLQ & "BF_TCOST, BF_UNITCOST, BF_PREMIUM,BF_PER, "
SQLQ = SQLQ & "BF_PPAMT, BF_MTHCCOST, BF_MTHECOST, BF_TAXBEN, "
If glbWFC Then
    If glbDivTranInPlant = "Y" Then 'Ticket #25275 Franks 04/02/2014
        SQLQ = SQLQ & "'Y', "
    Else
        SQLQ = SQLQ & "null, "
    End If
End If
SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, '"
SQLQ = SQLQ & Time$ & "' As AU_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' AS AU_LUSER, " & IIf(glbWFC, "'R'", "'A'") & " As AU_TYPE,'N' AS AU_UPLOAD FROM Term_HRBENFT "
SQLQ = SQLQ & "WHERE BF_EMPNBR =" & EEID&

gdbAdoIhr001.Execute SQLQ

RehBENEFITSAudit = True

Exit Function

RehBENEFITSAudit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehBENEFITSAudit", "Term_HRBENFT", "Insert")
Call RollBack '29July99 js

End Function

Private Function RehHREMPAudit(EEID&)
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim xProvNbr, xADD, xPROV
Dim Langs 'George Apr 4,2006 #10574
'Dim TIHR_DB As Database

On Error GoTo RehHREMPAudit_ERR

RehHREMPAudit = False

rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
rsTC.Open "HREMP", gdbAdoIhr001X, adOpenKeyset, , adCmdTableDirect
rsTC.Find "ED_EMPNBR = " & EEID&

OldDiv = ""
If rsTC.EOF Then
    MsgBox "SYSTEM ERROR - READING HREMP"
    Exit Function
End If

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_COMPNO") = "001"
If glbSamuel Then
    'Ticket #21791 Franks 04/02/2012. they don't want to create INEWEMP2 record
    'in Insync interface for Transfer In
    rsTA("AU_NEWEMP") = "N"
Else
    rsTA("AU_NEWEMP") = "Y"
End If
rsTA("AU_TYPE") = "R"
rsTA("AU_DIV") = clpDIV
rsTA("AU_DIVUPL") = clpDIV
If glbLinamar Then
    rsTA("AU_LOC") = rsTC("ED_LOC")
End If
If glbWFC Then
    rsTA("AU_LOC") = clpCode(7).Text
End If

rsTA("AU_EMPNBR") = rsTC("ED_EMPNBR")
rsTA("AU_TITLE") = rsTC("ED_TITLE")
rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
rsTA("AU_FNAME") = rsTC("ED_FNAME")
rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
rsTA("AU_CITY") = rsTC("ED_CITY")
rsTA("AU_PROV") = rsTC("ED_PROV")
rsTA("AU_PCODE") = rsTC("ED_PCODE")
rsTA("AU_PHONE") = rsTC("ED_PHONE")
rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
rsTA("AU_PROVRES") = rsTC("ED_PROVEMP")
rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
rsTA("AU_UIC") = rsTC("ED_UIC")
rsTA("AU_PENSION") = rsTC("ED_PENSION")
rsTA("AU_CPP") = rsTC("ED_CPP")
rsTA("AU_GROSSCD") = rsTC("ED_GROSSCD")
rsTA("AU_GARN") = rsTC("ED_GARN")
rsTA("AU_ELIGIBLE") = rsTC("ED_ELIGIBLE")
rsTA("AU_EARLYR") = rsTC("ED_EARLYR")
rsTA("AU_NORMALR") = rsTC("ED_NORMALR")
rsTA("AU_LATESTR") = rsTC("ED_LATESTR")
rsTA("AU_SIN") = rsTC("ED_SIN")
If glbWFC Then 'Ticket #23247 Franks 07/23/2013
    If Len(clpCode(12).Text) > 0 Then rsTA("AU_EMP") = clpCode(12).Text Else rsTA("AU_EMP") = rsTC("ED_EMP")
Else
    rsTA("AU_EMP") = rsTC("ED_EMP")
End If
rsTA("AU_PT") = rsTC("ED_PT")
rsTA("AU_PTUPL") = rsTC("ED_PT")
rsTA("AU_EMPTYPE") = rsTC("ED_EMPTYPE")
rsTA("AU_SEX") = rsTC("ED_SEX")
If rsTC("ED_SMOKER") <> 0 Then
    rsTA("AU_SMOKER") = "Yes"
Else
    rsTA("AU_SMOKER") = "No"
End If
rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
rsTA("AU_DEPTNO") = clpDept
rsTA("AU_DOB") = rsTC("ED_DOB")
'rsTA("AU_DOH") = rsTC("ED_DOH")
'rsTA("AU_SENDTE") = dlpDate(0)
'rsTA("AU_LTHIRE") = dlpDate(0) 'rsTC("ED_LTHIRE")
rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
rsTA("AU_BANK") = rsTC("ED_BANK")
rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
rsTA("AU_BANK2") = rsTC("ED_BANK2")
rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
rsTA("AU_PCDEPOSIT2") = rsTC("ED_PCDEPOSIT2")
rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
rsTA("AU_BANK3") = rsTC("ED_BANK3")
rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
rsTA("AU_SUPCODE") = rsTC("ED_SUPCODE")
rsTA("AU_DDI") = rsTC("ED_DDI")
rsTA("AU_WCB") = rsTC("ED_WCB")
If Len(clpCode(10).Text) > 0 Then rsTA("AU_ORG") = clpCode(10).Text Else rsTA("AU_ORG") = rsTC("ED_ORG")
rsTA("AU_UNION") = rsTC("ED_UNION")
rsTA("AU_TD1") = rsTC("ED_TD1")
rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
rsTA("AU_TD3") = rsTC("ED_TD3")
rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
rsTA("AU_ProvAmt") = rsTC("ED_PROVAMT") 'Ticket# 9843
rsTA("AU_VACPC") = rsTC("ED_VACPC")
rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
rsTA("AU_FDAY") = rsTC("ED_FDAY")
rsTA("AU_LDAY") = rsTC("ED_LDAY")
rsTA("AU_OMDAY") = rsTC("ED_OMERS")

rsTA("AU_INTEL") = rsTC("ED_INTEL")
'George Apr 4,2006 #10574
'rsTA("AU_LANG1") = rsTC("ED_LANG1")
'rsTA("AU_LANG2") = rsTC("ED_LANG2")
Langs = Split(getLanguage(rsTC("ED_EMPNBR")), "|")
If Langs(0) <> "NoLang1" Then rsTA("AU_LANG1") = Langs(0) '0 is for ED_Lang1
If Langs(1) <> "NoLang2" Then rsTA("AU_LANG2") = Langs(1) '1 is for ED_Lang2
'George Apr 4,2006 #10574

rsTA("AU_EMAIL") = rsTC("ED_EMAIL")
rsTA("AU_WITHSPOUSE") = rsTC("ED_WITHSPOUSE")
rsTA("AU_EXPYEAR") = rsTC("ED_EXPYEAR")

If glbLinamar Then
    rsTA("AU_SECTION") = clpCode(3)
    rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
    rsTA("AU_REGION") = clpCode(2)
    rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
    rsTA("AU_DOH") = rsTC("ED_DOH")
    rsTA("AU_SENDTE") = dlpDate(0)
    rsTA("AU_LTHIRE") = dlpDate(0)
ElseIf glbSamuel Then
    rsTA("AU_LOC") = clpCode(7).Text
    If Len(clpCode(0)) > 0 Then rsTA("AU_SECTION") = clpCode(0) Else rsTA("AU_SECTION") = Null
    If Len(clpCode(5)) > 0 Then rsTA("AU_ADMINBY") = clpCode(5) Else rsTA("AU_ADMINBY") = Null
    If Len(clpCode(6)) > 0 Then rsTA("AU_REGION") = clpCode(6) Else rsTA("AU_REGION") = Null
    If Len(txtPayrollID.Text) > 0 Then rsTA("AU_PAYROLL_ID") = txtPayrollID.Text
    If IsDate(dlpDate(3).Text) Then rsTA("AU_DOH") = CVDate((dlpDate(3).Text)) Else rsTA("AU_DOH") = rsTC("ED_DOH")
    If IsDate(dlpDate(4).Text) Then rsTA("AU_SENDTE") = CVDate((dlpDate(4).Text)) Else rsTA("AU_SENDTE") = Null
    If IsDate(dlpDate(5).Text) Then rsTA("AU_LTHIRE") = CVDate((dlpDate(5).Text)) Else rsTA("AU_LTHIRE") = Null
    'Ticket #22260 Franks 07/27/2012 - begin
    If Len(clpSalDist.Text) > 0 Then rsTA("AU_SALDIST") = clpSalDist.Text
    If Len(clpCode(11).Text) > 0 Then rsTA("AU_SUPCODE") = clpCode(11).Text
    If IsDate(dlpDate(7).Text) Then rsTA("AU_OMDAY") = CVDate((dlpDate(7).Text))
    If Len(clpVadim1.Text) > 0 Then rsTA("AU_VADIM1") = clpVadim1.Text
    If IsDate(dlpDate(8).Text) Then rsTA("AU_USRDAT1") = CVDate((dlpDate(8).Text))
    'Ticket #22260 Franks 07/27/2012 - end
Else 'WFC
    If Len(clpCode(0)) > 0 Then rsTA("AU_SECTION") = clpCode(0) Else rsTA("AU_SECTION") = Null
    If Len(clpCode(5)) > 0 Then rsTA("AU_ADMINBY") = clpCode(5) Else rsTA("AU_ADMINBY") = Null
    If Len(clpCode(6)) > 0 Then rsTA("AU_REGION") = clpCode(6) Else rsTA("AU_REGION") = Null
    If Len(txtPayrollID) > 0 Then rsTA("AU_PAYROLL_ID") = txtPayrollID
    If IsDate(dlpDate(3).Text) Then rsTA("AU_DOH") = CVDate((dlpDate(3).Text)) Else rsTA("AU_DOH") = rsTC("ED_DOH")
    If IsDate(dlpDate(4).Text) Then rsTA("AU_SENDTE") = CVDate((dlpDate(4).Text)) Else rsTA("AU_SENDTE") = Null
    If IsDate(dlpDate(5).Text) Then rsTA("AU_LTHIRE") = CVDate((dlpDate(5).Text)) Else rsTA("AU_LTHIRE") = Null
    If Len(clpVadim2.Text) > 0 Then rsTA("AU_VADIM2") = ((clpVadim2.Text))
    'Ticket #22448 Franks 10/31/2012 - begin
    If Len(txtUserNum1.Text) > 0 Then rsTA("AU_USER_NUM1") = txtUserNum1.Text
    If Len(txtUserText2.Text) > 0 Then rsTA("AU_USER_TEXT2") = txtUserText2.Text
    'Ticket #22448 Franks 10/31/2012 - end
End If
rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
rsTA("AU_SSN") = rsTC("ED_SSN")
rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")

If glbLinamar Then
    rsTA("AU_ExtrAnn") = rsTC("ED_ExtrAnn")
    rsTA("AU_QTBTORRSP") = rsTC("ED_QTBTORRSP")
    rsTA("AU_LOCKER") = rsTC("ED_LOCKER")
    rsTA("AU_COMBINATION") = rsTC("ED_COMBINATION")
End If

If Len(clpGLNum.Text) > 0 Then rsTA("AU_DEPT_GL") = clpGLNum.Text
rsTA("AU_HOMEOPRTNBR") = clpHOME(1)
rsTA("AU_HOMELINE") = clpHOME(2)
rsTA("AU_HOMEWRKCNT") = clpHOME(3)
rsTA("AU_HOMESHIFT") = clpHOME(4)
rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"

rsTA("AU_TIDIV") = clpDIV & ""
rsTA("AU_TIREASON_TABL") = "SDJC"
rsTA("AU_TIREASON") = clpCode(1)
rsTA("AU_TIDATE") = dlpDate(0)
rsTA("AU_TIEMPNBR") = rsTC("ED_EMPNBR")

rsTA("AU_TODIV") = Right(Trim(Str(rsORG("TL_EMPNBR"))), 3)
rsTA("AU_TOEMPNBR") = rsORG("TL_EMPNBR")

rsTA("AU_TCOMPLETE") = "Y"
If glbWFC And glbDivTranInPlant = "Y" Then 'Ticket #25275 Franks 04/02/2014
    rsTA("AU_VSTEP") = "Y"
End If
rsTA.Update
rsTA.Close

Dim xSalCD, rsTB As New ADODB.Recordset
rsTB.Open "SELECT SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EEID&, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    xSalCD = ""
Else
    xSalCD = rsTB!SH_SALCD
End If
Dim xKey
xKey = "E" & rsTC!ED_EMPNBR
rsTA.Open "LN_TRALOG", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rsTA.AddNew

rsTA!TL_COMPNO = "001"
rsTA!TL_EMPNBR = rsTC!ED_EMPNBR
rsTA!TL_SURNAME = rsTC("ED_SURNAME")
rsTA!TL_FNAME = rsTC("ED_FNAME")
If glbLinamar Then
    rsTA!TL_DOH = rsTC!ED_DOH
Else
    If IsDate(dlpDate(3).Text) Then
        rsTA!TL_DOH = CVDate((dlpDate(3).Text))
    Else
        rsTA!TL_DOH = rsTC!ED_DOH
    End If
End If
rsTA!TL_JOB = clpJob

rsTA!TL_TYPE = "TIN"

rsTA!TL_OLDDIV = rsORG!TL_OLDDIV
If Not IsNull(rsORG!TL_OLDDIV) Then
    OldDiv = rsORG!TL_OLDDIV
End If
rsTA!TL_OLDEMPNBR = rsORG!TL_OLDEMPNBR
rsTA!TL_OLDDIVEDATE = rsORG!TL_OLDDIVEDATE

rsTA!TL_NEWDIV = clpDIV & ""
rsTA!TL_NEWEMPNBR = rsTC("ED_EMPNBR")
rsTA!TL_NEWDIVEDATE = dlpDate(0)



rsTA!TL_TOREASON_TABL = "TERM"
rsTA!TL_TIREASON_TABL = "SDJC"
rsTA!TL_TIREASON = clpCode(1)
rsTA!TL_SALCD = xSalCD
rsTA!TL_TERM_SEQ = rsORG!TL_TERM_SEQ
rsTA!TL_TCOMPLETE = "Y"

rsTA!TL_KEY = xKey
rsTA!TL_CURRENTDIV = clpDIV & ""

rsTA("TL_LDATE") = Format(Now, "SHORT DATE")
rsTA("TL_LUSER") = glbUserID
rsTA("TL_LTIME") = Time$
rsTA.Update

rsTA.Close
rsTA.Open "SELECT TL_KEY,TL_CURRENTDIV FROM LN_TRALOG WHERE TL_KEY='T" & rsORG!TL_TERM_SEQ & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Do Until rsTA.EOF
    rsTA!TL_KEY = xKey
    rsTA!TL_CURRENTDIV = clpDIV & ""
    rsTA.Update
    rsTA.MoveNext
Loop



RehHREMPAudit = True

Exit Function

RehHREMPAudit_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function

Private Function RehJOBAudit(EEID&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String

RehJOBAudit = False

On Error GoTo RehJOBAudit_Err

SQLQ = "INSERT INTO HRAUDIT ( AU_COMPNO, AU_EMPNBR, AU_WHRS, "
SQLQ = SQLQ & "AU_DHRS, AU_PHRS, AU_SJDATE, "
If glbLinamar Then SQLQ = SQLQ & "AU_LEADHAND,AU_LABOURCD,"
If glbWFC Then SQLQ = SQLQ & "AU_PAYROLL_ID,AU_VSTEP," 'Ticket #25275 Franks 04/02/2014 added AU_VSTEP
If glbSamuel Then SQLQ = SQLQ & "AU_ADMINBY,"
SQLQ = SQLQ & "AU_JOB, AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE,AU_UPLOAD ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)

SQLQ = SQLQ & " VALUES( '001', " & EEID & " , " & medHours(1) & ", "
SQLQ = SQLQ & medHours(0) & "," & medHours(2) & ", "
If glbWFC Then 'Ticket #24652 Franks 11/14/2013
    SQLQ = SQLQ & Date_SQL(dlpDate(1)) & " , "
Else
    SQLQ = SQLQ & Date_SQL(dlpDate(0)) & " , "
End If
'If glbLinamar Then SQLQ = SQLQ & "'','" & clpCode(4) & "',"
If glbLinamar Then SQLQ = SQLQ & "'','" & txtLabCode.Text & "',"  'Ticket #29946 Franks 03/15/2017
If glbWFC Then 'Ticket #12366
    If Len(txtPayrollID) > 0 Then
        SQLQ = SQLQ & "'" & txtPayrollID & "', "
    Else
        SQLQ = SQLQ & "null, "
    End If
    If glbDivTranInPlant = "Y" Then 'Ticket #25275 Franks 04/02/2014
        SQLQ = SQLQ & "'Y', "
    Else
        SQLQ = SQLQ & "null, "
    End If
End If
If glbSamuel Then SQLQ = SQLQ & "'" & clpCode(5).Text & "',"
SQLQ = SQLQ & "'" & clpJob & "', "
SQLQ = SQLQ & Date_SQL(Date) & ", "
SQLQ = SQLQ & "'" & Time$ & "', "
SQLQ = SQLQ & "'" & glbUserID & "', " & IIf(glbWFC, "'R'", "'A'") & ",'N' )"

gdbAdoIhr001.Execute SQLQ

RehJOBAudit = True

Exit Function

RehJOBAudit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehJOBAudit", "Term_JOB_HISTORY", "Insert")
Call RollBack '29July99 js

End Function

Private Function RehSALARYAudit(EEID&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String

RehSALARYAudit = False

On Error GoTo RehSALARYAudit_Err

SQLQ = "INSERT INTO HRAUDIT ( AU_COMPNO, AU_EMPNBR, AU_SALCD, "
If glbWFC Then SQLQ = SQLQ & "AU_PAYROLL_ID,AU_VSTEP,"
If glbSamuel Then SQLQ = SQLQ & "AU_ADMINBY,"
SQLQ = SQLQ & "AU_PAYP, AU_SEDATE, AU_SNDATE, "
SQLQ = SQLQ & "AU_SALARY, AU_SJDATE, "
SQLQ = SQLQ & "AU_JOB, AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE,AU_UPLOAD ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)

SQLQ = SQLQ & " VALUES( '001', " & EEID & " , '" & lblSalCode & "', "
If glbWFC Then
    If Len(txtPayrollID) > 0 Then
        SQLQ = SQLQ & "'" & txtPayrollID & "', "
    Else
        SQLQ = SQLQ & "null, "
    End If
    If glbDivTranInPlant = "Y" Then 'Ticket #25275 Franks 04/02/2014
        SQLQ = SQLQ & "'Y', "
    Else
        SQLQ = SQLQ & "null, "
    End If
End If
If glbSamuel Then SQLQ = SQLQ & "'" & clpCode(5).Text & "',"
If glbSamuel Then
    SQLQ = SQLQ & "'" & clpCode(9).Text & "', " & Date_SQL(dlpDate(2)) & ", Null, "
Else
    SQLQ = SQLQ & "'', " & Date_SQL(dlpDate(0)) & ", " & Date_SQL(dlpDate(2)) & ", "
End If
SQLQ = SQLQ & medSalary & ", " & Date_SQL(dlpDate(0)) & ", '" & clpJob & "', "
If glbWFC Then 'Ticket #24652 Franks 11/14/2013
    SQLQ = SQLQ & Date_SQL(dlpDate(1)) & " , "
Else
    SQLQ = SQLQ & Date_SQL(dlpDate(0)) & ", "
End If
SQLQ = SQLQ & "'" & Time$ & "', "
SQLQ = SQLQ & "'" & glbUserID & "', " & IIf(glbWFC, "'R'", "'A'") & ",'N') "

gdbAdoIhr001.Execute SQLQ
RehSALARYAudit = True


Exit Function

RehSALARYAudit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehSALARYAudit", "Term_SALARY_HISTORY", "Insert")
Call RollBack '29July99 js

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



Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Left = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
glbOnTop = ""
End Sub

Private Sub medHours_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub medSalary_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub lblEmpNo_Change()
Call EmpNoExist(getEmpnbr(lblEmpNo))
End Sub

Private Sub Job_Desc()
Dim SQLQ As String
Dim X%
Dim rsJOB As New ADODB.Recordset
On Error GoTo Jobd_Err

If Len(clpJob.Text) > 0 Then
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & CStr(clpJob.Text) & "'"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        If Not IsNull(rsJOB("JB_S" & X%)) Then JobSnap_PayScale(X) = Round2DEC(rsJOB("JB_S" & X%))
    Next
    If Not IsNull(rsJOB("JB_SALCD")) Then JobSnap_Salary_Code$ = rsJOB("JB_SALCD")
    If Not IsNull(rsJOB("JB_MIDPOINT")) Then JobSnap_MidPoint! = rsJOB("JB_MIDPOINT")
End If

Exit Sub

Jobd_Err:
If Err = 94 Then
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "JOBS", "SELECT")
Call RollBack '26July99 js

End Sub

Private Sub Dept_GL()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String
Dim rsDEPT As New ADODB.Recordset
On Error GoTo Dept_GL_Err

If Len(clpDept.Text) > 0 Then
    rsDEPT.Open "SELECT DF_GLNO FROM HRDEPT WHERE DF_NBR='" & clpDept.Text & "'", gdbAdoIhr001
    If Not rsDEPT.EOF Then
        RGLNum = rsDEPT("DF_GLNO")
        If RDept <> clpDept Then
            If IsNull(RGLNum) Then
                RGLNum = ""
            Else
                Msg$ = "Do you want the associated G/L #?"
                Title$ = "info:HR"
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDYES Then clpGLNum.Text = RGLNum
            End If
            RDept = clpDept.Text
        End If
    End If
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
     clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Dept Snap", "DEPT", "SELECT")
Call RollBack '21June99 js
End Sub




Private Sub clpDIV_Change()
    If Len(clpDIV) > 0 And clpDIV.Caption <> "Unassigned" And InStr(lblEmpNo, "-") > 3 Then
        lblEmpNo.Caption = clpDIV & Mid(lblEmpNo, InStr(lblEmpNo, "-"))
        If glbLinamar Then
        Call set_CountrySIN(clpDIV)
        End If
    End If

End Sub

Private Sub CountEmpNbr()

If glbSamuel Then
    lblEmpNo.Visible = False
    Exit Sub
Else
    lblEmpNo.Visible = True
End If
If glbLinamar Then
    If Len(clpDIV) = 3 And Val(txtEmpID) > 0 Then
        lblEmpNo = Format(clpDIV, "000") & "-" & Val(txtEmpID)
    Else
        lblEmpNo = ""
    End If
End If
If glbWFC Then
    If Len(clpDIV) = 4 And Val(txtEmpID) > 0 Then
        lblEmpNo = clpDIV & (txtEmpID)
    Else
        lblEmpNo = ""
    End If
End If
End Sub

Private Sub txtEmpID_Change()
Call CountEmpNbr
End Sub
Private Sub txtEmpID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
Private Sub UPDMOD()
Dim X%
For X% = 0 To 2
    dlpDate(X%).Enabled = False
     clpCode(X% + 1).Enabled = False
     clpHOME(X% + 1).Enabled = False
    medHours(X%).Enabled = False
Next
 clpDIV.Enabled = False
txtEmpID.Enabled = False
 clpDept.Enabled = False
 clpGLNum.Enabled = False
 clpHOME(4).Enabled = False
 clpJob.Enabled = False
medSalary.Enabled = False
End Sub

Private Function Upd_Related_Salary(EID&)
Dim SQLQ As String, Msg As String
Dim dynHRSALHIS As New ADODB.Recordset
Dim JobCode$, PositionStartDat, JobReason$
Dim HoursPerWeek!
Dim X!, cX$
Dim SH_SALARY@, SH_SALCD$, SH_EDATE, SH_PAYP$, SH_NEXTDAT As Variant
Dim SHisDate As Variant
Dim AnnualSalary As Double, Compa!, SalaryGrade$
Dim xPosEarly

On Error GoTo UpRel_Err

JobCode$ = clpJob.Text
If glbWFC Then 'Ticket #24652 Franks 11/14/2013
    PositionStartDat = CVDate(dlpDate(1).Text)
Else
    PositionStartDat = CVDate(dlpDate(0).Text)
End If

If glbLinamar Then
    JobReason$ = "TRAN"
Else 'WFC
    JobReason$ = clpCode(8) '"TRNI"
End If
SH_SALARY@ = medSalary
SH_NEXTDAT = dlpDate(2)
SH_SALCD$ = lblSalCode
HoursPerWeek! = Val(medHours(1))

SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & EID&
SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL, "DESC", "")
dynHRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not dynHRSALHIS.EOF Then
    If IsDate(dynHRSALHIS("SH_EDATE")) Then
        SHisDate = CVDate(dynHRSALHIS("SH_EDATE"))
    Else
        SHisDate = PositionStartDat
    End If
End If

xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0

If Not (dynHRSALHIS.BOF And dynHRSALHIS.EOF) Then
    Do Until dynHRSALHIS.EOF
        dynHRSALHIS("SH_CURRENT") = False
        dynHRSALHIS.Update
        dynHRSALHIS.MoveNext
    Loop
    dynHRSALHIS.MoveFirst
End If

'SET COMPA RATIO
'================

If JobSnap_Salary_Code$ = "A" Then
    If SH_SALCD$ = "H" Then
        AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52
    ElseIf SH_SALCD$ = "A" Then
        AnnualSalary = SH_SALARY@
    ElseIf SH_SALCD$ = "M" Then
        AnnualSalary = SH_SALARY@ * 12
    'added by Bryan 27/Sep/05 Ticket#9354
    ElseIf SH_SALCD$ = "D" Then
        If GetLeapYear(Year(Date)) Then
            AnnualSalary = SH_SALARY@ * 366
        Else
            AnnualSalary = SH_SALARY@ * 365
        End If
    End If
Else
    If SH_SALCD$ = "A" Then
        If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 52
    ElseIf SH_SALCD$ = "A" Then
        AnnualSalary = SH_SALARY@
    ElseIf SH_SALCD$ = "M" Then
        If HoursPerWeek! = 0 Then
            AnnualSalary = 0
        Else
            AnnualSalary = (SH_SALARY@ * 12 / HoursPerWeek!) / 52
        End If
    'added by Bryan 26/Sep/05 Ticket#9354
    ElseIf SH_SALCD$ = "D" Then
        If HoursPerWeek! = 0 Then
            AnnualSalary = 0
        Else
            If GetLeapYear(Year(Date)) Then
                AnnualSalary = (SH_SALARY@ * 366 / HoursPerWeek!) / 52
            Else
                AnnualSalary = (SH_SALARY@ * 365 / HoursPerWeek!) / 52
            End If
        End If
    End If
End If
 ' set COMPA RATIO
If glbWFC Then 'Ticket #24936 Franks 02/05/2014
    Compa! = Get_WFC_COMPA_FromMaster(clpCode(10).Text, clpJob.Text, medSalary.Text, clpCode(0).Text, cmbMarketLine.Text, txtFiscalYear.Text)
    '(xUnion, xJob, xSalary, fglbSection, xMarketLine, xFiscalYear) 'Ticket #25045 Franks 02/05/2014
Else
    If JobSnap_PayScale(JobSnap_MidPoint!) <> 0 And AnnualSalary <> 0 Then
        Compa! = (AnnualSalary / JobSnap_PayScale(JobSnap_MidPoint!)) * 100
    Else
        Compa! = 0
    End If
End If

If Compa! > 999.99 Then
    Compa! = 999.99
End If

'Determine Pay Scale individual fits into
'==========================================
SalaryGrade$ = "00"
'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
'For X! = 1 To 11
'For X! = 1 To 15
For X! = 1 To 20
    If AnnualSalary >= JobSnap_PayScale(X) And JobSnap_PayScale(X) > 0 Then
      cX$ = CStr(X)
      If X! <= 9 Then cX$ = "0" & cX$
      SalaryGrade$ = cX$
    End If
Next X!

'NOW UPDATE SALARY HISTORY TABLE  - only if new record do we add record
'================================
dynHRSALHIS.AddNew

dynHRSALHIS("SH_COMPNO") = "001"
dynHRSALHIS("SH_EMPNBR") = EID&
dynHRSALHIS("SH_CURRENT") = True
dynHRSALHIS("SH_SDATE") = PositionStartDat
If glbSamuel Then
    dynHRSALHIS("SH_EDATE") = dlpDate(2).Text
Else
    dynHRSALHIS("SH_EDATE") = IIf(xPosEarly, SHisDate, CVDate(PositionStartDat))
End If
dynHRSALHIS("SH_SALARY") = SH_SALARY@
dynHRSALHIS("SH_SALCD") = SH_SALCD$
dynHRSALHIS("SH_JOB") = JobCode$
dynHRSALHIS("SH_JOB_ID") = fglbJobID&
dynHRSALHIS("SH_PAYP_TABLE") = "SDPP"
If Len(clpCode(9).Text) > 0 Then dynHRSALHIS("SH_PAYP") = clpCode(9).Text 'Ticket #15818
If glbSamuel Then 'use this date as salary effective date
Else
If IsDate(SH_NEXTDAT) Then dynHRSALHIS("SH_NEXTDAT") = SH_NEXTDAT
End If
dynHRSALHIS("SH_WHRS") = HoursPerWeek!
dynHRSALHIS("SH_SREAS_TABLE") = "SDRC"
dynHRSALHIS("SH_SREAS1") = JobReason$
dynHRSALHIS("SH_COMPA") = Round(Compa!, 2)
dynHRSALHIS("SH_GRADE") = Format(SalaryGrade$, "00")
dynHRSALHIS("SH_TRANSDATE") = Date
dynHRSALHIS("SH_LDATE") = Date
dynHRSALHIS("SH_LTIME") = Time$
dynHRSALHIS("SH_LUSER") = glbUserID
If glbWFC And txtFiscalYear.Visible = True Then
    If IsNumeric(txtFiscalYear.Text) Then 'Ticket #17742
        dynHRSALHIS("SH_FISCALYEAR") = txtFiscalYear.Text
    End If
    dynHRSALHIS("SH_MARKETLINE") = txtMarketLine.Text
    dynHRSALHIS("SH_SECTION") = clpCode(0).Text
    dynHRSALHIS("SH_BAND") = fglbBAND
End If
If glbSamuel Then
    dynHRSALHIS("SH_GRID") = clpGrid.Text
End If
If glbWFC Then 'Ticket #29888 Franks 05/03/2017 - • Transfer In did not update currency
    dynHRSALHIS("SH_CURRENCYINDI") = getWFCCurrencyIndi(clpCode(0).Text)
End If
dynHRSALHIS.Update
dynHRSALHIS.Close
Call updFollow(EID&, "S")
If glbWFC Then
    'Ticket #23247 Franks 09/16/2013 - do this later
Else
    Call updBenefitForSalDEPN(EID&)
End If

Exit Function

UpRel_Err:
If Err = 3021 Then
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HRSAL/PERF", "INSERT")
Call RollBack '26July99 js

End Function



Private Sub Upd_Related_Job(EID&)
Dim SQLQ As String, Msg As String
Dim dynHRJOBHIS As New ADODB.Recordset
Dim HRJH_Snap As New ADODB.Recordset
Dim JobCode$, PositionStartDat, JobReason$
Dim HoursPerDay!, HoursPerWeek!, HoursPerPayPeriod!
Dim xRet4
On Error GoTo UpRel_Err

JobCode$ = clpJob.Text
If glbWFC Then 'Ticket #24652 Franks 11/14/2013
    PositionStartDat = CVDate(dlpDate(1).Text)
Else
    PositionStartDat = CVDate(dlpDate(0).Text)
End If

If glbLinamar Then
    JobReason$ = "TRAN"
Else 'wfc
    JobReason$ = clpCode(8) '"TRNI"
End If
HoursPerDay! = Val(medHours(0))
HoursPerWeek! = Val(medHours(1))
HoursPerPayPeriod! = Val(medHours(2))

SQLQ = "SELECT * FROM HR_JOB_HISTORY"
SQLQ = SQLQ & " WHERE JH_EMPNBR = " & EID&
SQLQ = SQLQ & " ORDER BY JH_SDATE DESC "
dynHRJOBHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not (dynHRJOBHIS.BOF And dynHRJOBHIS.EOF) Then
    dynHRJOBHIS("JH_CURRENT") = False
    dynHRJOBHIS.Update
End If
dynHRJOBHIS.AddNew
dynHRJOBHIS("JH_COMPNO") = "001"
dynHRJOBHIS("JH_EMPNBR") = EID&
dynHRJOBHIS("JH_CURRENT") = True
dynHRJOBHIS("JH_SDATE") = CVDate(PositionStartDat)
dynHRJOBHIS("JH_JOB") = JobCode$
dynHRJOBHIS("JH_DHRS") = HoursPerDay!
dynHRJOBHIS("JH_WHRS") = HoursPerWeek!
dynHRJOBHIS("JH_PHRS") = HoursPerPayPeriod!
dynHRJOBHIS("JH_ENDREAS_TABL") = "SDRC"
dynHRJOBHIS("JH_JREASON") = JobReason$
If glbLinamar Then
    'dynHRJOBHIS("JH_SHIFT") = txtShift
    dynHRJOBHIS("JH_SHIFT") = Left((clpDIV.Text & clpCode(13).Text), 20)  'Ticket #29414 Franks 11/03/2016
    'dynHRJOBHIS("JH_LABOURCD") = clpCode(4)
    dynHRJOBHIS("JH_LABOURCD") = txtLabCode.Text  'Ticket #29946 Franks 03/15/2017
    dynHRJOBHIS("JH_DIV") = clpDIV.Text  'Ticket# 8293
    If elpReptAuthShow(0).Visible Then 'Ticket #29414 Franks 11/03/2016
        If Len(elpReptAuthShow(0).Text) > 0 Then
            dynHRJOBHIS("JH_REPTAU") = getEmpnbr(elpReptAuthShow(0).Text)
        Else
            dynHRJOBHIS("JH_REPTAU") = Null
        End If
    End If
End If
If glbSamuel Then
    dynHRJOBHIS("JH_GRID") = clpGrid.Text
    'Ticket #21791 Franks 04/09/2012
    If IsNumeric(elpReptAuthShow(0).Text) Then dynHRJOBHIS("JH_REPTAU") = elpReptAuthShow(0).Text Else dynHRJOBHIS("JH_REPTAU") = Null
    If IsNumeric(elpReptAuthShow(1).Text) Then dynHRJOBHIS("JH_REPTAU2") = elpReptAuthShow(1).Text Else dynHRJOBHIS("JH_REPTAU2") = Null
    If IsNumeric(elpReptAuthShow(2).Text) Then dynHRJOBHIS("JH_REPTAU3") = elpReptAuthShow(2).Text Else dynHRJOBHIS("JH_REPTAU3") = Null
    If IsNumeric(elpReptAuthShow(3).Text) Then dynHRJOBHIS("JH_REPTAU4") = elpReptAuthShow(3).Text Else dynHRJOBHIS("JH_REPTAU4") = Null
    If chkProSha.Value Then dynHRJOBHIS("JH_PROFIT_SHARING") = 1 Else dynHRJOBHIS("JH_PROFIT_SHARING") = 0
End If
If glbWFC And elpReptAuthShow(0).Visible Then  'Ticket #24936 Franks 02/05/2014
    If IsNumeric(elpReptAuthShow(0).Text) Then
        dynHRJOBHIS("JH_REPTAU") = elpReptAuthShow(0).Text
        If IsDate(dlpDate(1).Text) Then
            dynHRJOBHIS("JH_EDATEREPT1") = CVDate(dlpDate(1).Text)
        End If
    Else
        dynHRJOBHIS("JH_REPTAU") = Null
    End If
End If
dynHRJOBHIS("JH_LDATE") = Date
dynHRJOBHIS("JH_LTIME") = Time$
dynHRJOBHIS("JH_LUSER") = glbUserID

If glbWFC And glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/18/2014
    '-------------Position begin -------------
    SQLQ = "Select * from TERM_JOB_HISTORY WHERE TERM_SEQ=" & fglbTERM_Seq & " AND JH_CURRENT <>0"
    HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not HRJH_Snap.EOF Then
        dynHRJOBHIS("JH_SHIFT") = HRJH_Snap("JH_SHIFT")
        dynHRJOBHIS("JH_FTENUM") = HRJH_Snap("JH_FTENUM")
        dynHRJOBHIS("JH_FTEHRS") = HRJH_Snap("JH_FTEHRS")
        dynHRJOBHIS("JH_REPTAU2") = HRJH_Snap("JH_REPTAU2")
        dynHRJOBHIS("JH_REPTAU3") = HRJH_Snap("JH_REPTAU3")
        dynHRJOBHIS("JH_REPTAU4") = HRJH_Snap("JH_REPTAU4")
    End If
    HRJH_Snap.Close
    '-------------Position end -------------
End If
If glbWFC Then 'Ticket #28254 Franks 03/22/2016
    xRet4 = getWFCRA4(clpDIV.Text)
    If Len(xRet4) = 0 Then
        dynHRJOBHIS("JH_REPTAU4") = Null
    Else
        dynHRJOBHIS("JH_REPTAU4") = xRet4
    End If
End If

dynHRJOBHIS.Update
fglbJobID& = dynHRJOBHIS("JH_ID")
dynHRJOBHIS.Close

Exit Sub

UpRel_Err:
If Err = 3021 Then
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "JOB HISTORY", "HRJOB/PERF", "INSERT")
Call RollBack '26July99 js

End Sub



Private Sub Upd_Related_EMP(EID&)
Dim SQLQ As String, Msg As String
Dim dynHREMP As New ADODB.Recordset
Dim StartDat, HoursPerDay!
Dim xEmplCountry

On Error GoTo UpRel_Err
StartDat = CVDate(dlpDate(0).Text)
HoursPerDay! = Val(medHours(0))

SQLQ = "SELECT * FROM HREMP"
SQLQ = SQLQ & " WHERE ED_EMPNBR = " & EID&
dynHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If glbLinamar Then
    'dynHREMP("ED_COUNTRY") = "CANADA" 'Ticket #22819 Franks 11/14/2012
    'dynHREMP("ED_PROV") = "ON"
    'If medSIN.Visible Then dynHREMP("ED_SIN") = medSIN
    dynHREMP("ED_GLNO") = clpGLNum
    dynHREMP("ED_REGION") = clpDIV & clpCode(2)
    dynHREMP("ED_SECTION") = clpDIV & clpCode(3)
    dynHREMP("ED_OMERS") = Null
    dynHREMP("ED_HOMEOPRTNBR") = clpDIV & clpHOME(1)
    dynHREMP("ED_HOMELINE") = clpDIV & clpHOME(2)
    dynHREMP("ED_HOMEWRKCNT") = clpHOME(3)
    dynHREMP("ED_HOMESHIFT") = clpHOME(4)
    dynHREMP("ED_DEPTEDATE") = StartDat
    dynHREMP("ED_LTHIRE") = StartDat
    dynHREMP("ED_SENDTE") = StartDat
ElseIf glbSamuel Then
    dynHREMP("ED_ADMINBY") = clpCode(5).Text
    dynHREMP("ED_SECTION") = clpCode(0).Text
    dynHREMP("ED_LOC") = clpCode(7).Text
    dynHREMP("ED_REGION") = clpCode(6)
    If Len(clpGLNum.Text) > 0 Then
        dynHREMP("ED_GLNO") = clpGLNum.Text  'txtEmpID
    Else
        dynHREMP("ED_GLNO") = Null
    End If
    If Len(clpCode(10).Text) > 0 Then
        dynHREMP("ED_ORG") = clpCode(10).Text
    Else
        dynHREMP("ED_ORG") = Null
    End If
    If Len(txtPayrollID) > 0 Then
        dynHREMP("ED_PAYROLL_ID") = txtPayrollID.Text  'txtEmpID
    Else
        dynHREMP("ED_PAYROLL_ID") = Null
    End If
    dynHREMP("ED_DIVEDATE") = StartDat
    If Len(dlpDate(3).Text) > 0 Then dynHREMP("ED_DOH") = CVDate(dlpDate(3).Text)
    If IsDate((dlpDate(4).Text)) Then
        dynHREMP("ED_SENDTE") = CVDate(dlpDate(4).Text)
    Else
        dynHREMP("ED_SENDTE") = Null
    End If
    If IsDate((dlpDate(5).Text)) Then
        dynHREMP("ED_LTHIRE") = CVDate(dlpDate(5).Text)
    Else
        dynHREMP("ED_LTHIRE") = Null
    End If
    'Ticket #22260 Franks 07/27/2012 - begin
    If Len(clpBGroup.Text) > 0 Then dynHREMP("ED_BENEFIT_GROUP") = clpBGroup.Text Else dynHREMP("ED_BENEFIT_GROUP") = Null
    If Len(clpSalDist.Text) > 0 Then dynHREMP("ED_SALDIST") = clpSalDist.Text Else dynHREMP("ED_SALDIST") = Null
    If Len(clpCode(11).Text) > 0 Then dynHREMP("ED_SUPCODE") = clpCode(11).Text Else dynHREMP("ED_SUPCODE") = Null
    If IsDate((dlpDate(7).Text)) Then dynHREMP("ED_OMERS") = dlpDate(7).Text Else dynHREMP("ED_OMERS") = Null
    If Len(clpVadim1.Text) > 0 Then dynHREMP("ED_VADIM1") = clpVadim1.Text Else dynHREMP("ED_VADIM1") = Null
    If IsDate((dlpDate(8).Text)) Then dynHREMP("ED_USRDAT1") = dlpDate(8).Text Else dynHREMP("ED_USRDAT1") = Null
    'Ticket #22260 Franks 07/27/2012 - end
Else 'glbWFC
    dynHREMP("ED_SECTION") = clpCode(0)
    dynHREMP("ED_LOC") = clpCode(7)
    dynHREMP("ED_ADMINBY") = clpCode(5)
    dynHREMP("ED_REGION") = clpCode(6)
    If Len(txtPayrollID) > 0 Then
        dynHREMP("ED_PAYROLL_ID") = txtPayrollID 'txtEmpID
    'Else
    '    dynHREMP("ED_PAYROLL_ID") = Null
    End If
    dynHREMP("ED_OMERS") = Null
    dynHREMP("ED_DIVEDATE") = StartDat
    dynHREMP("ED_DOH") = CVDate(dlpDate(3).Text)
    If IsDate((dlpDate(4).Text)) Then
        dynHREMP("ED_SENDTE") = CVDate(dlpDate(4).Text)
    'Else
    '    dynHREMP("ED_SENDTE") = Null
    End If
    If IsDate((dlpDate(5).Text)) Then
        dynHREMP("ED_LTHIRE") = CVDate(dlpDate(5).Text)
    'Else
    '    dynHREMP("ED_LTHIRE") = Null
    End If
    'Ticket #16748
    If Len(clpVadim2.Text) > 0 Then dynHREMP("ED_VADIM2") = clpVadim2.Text
    If Len(clpGLNum.Text) > 0 Then dynHREMP("ED_GLNO") = clpGLNum.Text
End If
'Ticket #16395 Pension
If glbWFC Then
    If Len(clpCode(12).Text) > 0 Then dynHREMP("ED_EMP") = clpCode(12).Text 'Ticket #23247 Franks 07/23/2013
    If Len(clpCode(10).Text) > 0 Then
        dynHREMP("ED_ORG") = clpCode(10).Text
    End If
    If Len(txtEmpType.Text) > 0 Then
        dynHREMP("ED_EMPTYPE") = txtEmpType.Text
        If txtEmpType.Text = "Y" Then
            If IsNull(dynHREMP("ED_ELIGIBLE")) Then
                dynHREMP("ED_ELIGIBLE") = dlpDate(1).Text
            End If
        End If
    End If
    'Ticket #18654
    If Len(clpBGroup.Text) > 0 Then dynHREMP("ED_BENEFIT_GROUP") = clpBGroup.Text Else dynHREMP("ED_BENEFIT_GROUP") = Null
    'Ticket #22448 Franks 10/31/2012 - begin
    If Len(txtUserNum1.Text) > 0 Then dynHREMP("ED_USER_NUM1") = txtUserNum1.Text Else dynHREMP("ED_USER_NUM1") = Null
    If Len(txtUserText2.Text) > 0 Then
        dynHREMP("ED_USER_TEXT2") = txtUserText2.Text
    Else
        dynHREMP("ED_USER_TEXT2") = Null
        dynHREMP("ED_USER_TEXT1") = Null 'Ticket #24620 Franks 12/03/2013
    End If
    'Ticket #22448 Franks 10/31/2012 - end
    'If Len(clpVadim1.Text) > 0 Then dynHREMP("ED_VADIM1") = clpVadim1.Text Else dynHREMP("ED_VADIM1") = Null  'Ticket #23247 Franks 09/13/2013
    
    'Ticket #19678 Franks 01/24/2011
    'On Transfer In:   If Eligible for Pension equals "Y",
    '(Plant Code equals "TILB" and Union Code is "C127") or (Plant Code equals "WHBY" and Union Code is "C222") the Hire Code equals "Y".
    If txtEmpType.Text = "Y" Then
        If (clpCode(0).Text = "TILB" And clpCode(10).Text = "C127") Or (clpCode(0).Text = "WHBY" And clpCode(10).Text = "C222") Then
            dynHREMP("ED_HIRECODE") = "Y"
        End If
    End If
    
    'Ticket #24451 Franks 10/17/2013
    xEmplCountry = GetCountryFromDiv(clpDIV.Text)
    If Len(xEmplCountry) > 0 Then
        dynHREMP("ED_WORKCOUNTRY") = xEmplCountry
    End If
    
    If Len(clpCode(14).Text) > 0 Then 'Ticket #30376 Franks 07/17/2017
        dynHREMP("ED_ORGT1") = clpCode(14).Text
    End If
    
    'Ticket #19955 Franks 03/07/2011
    If dlpDate(6).Visible Then
        If clpCode(12).Text = "COOP" Or clpCode(12).Text = "STUD" Then
            'Ticket #25352 Franks 04/16/2014 -
            '"   If Employment Status = COOP or STUD, no NGS Start Date should transfer over
            xWFC_NewNGSStart = ""
        Else
            xWFC_NewNGSStart = dlpDate(6).Text
        End If
        If IsDate(xWFC_OldNGSStart) Then xWFC_OldNGSStart = CVDate(xWFC_OldNGSStart) Else xWFC_OldNGSStart = ""
        If IsDate(xWFC_NewNGSStart) Then xWFC_NewNGSStart = CVDate(xWFC_NewNGSStart) Else xWFC_NewNGSStart = ""
        If Not xWFC_OldNGSStart = xWFC_NewNGSStart Then
            If IsDate(xWFC_NewNGSStart) Then
                Call Upt_EmpOtherByField(EID&, "ER_OTHERDATE1", xWFC_NewNGSStart)
            Else
                Call Upt_EmpOtherByField(EID&, "ER_OTHERDATE1", Null)
            End If
        End If
    End If
End If
dynHREMP("ED_DEPTNO") = clpDept
dynHREMP("ED_DIV") = clpDIV
dynHREMP("ED_DHRS") = HoursPerDay!
dynHREMP("ED_LDATE") = Date
dynHREMP("ED_LTIME") = Time$
dynHREMP("ED_LUSER") = glbUserID
dynHREMP.Update
dynHREMP.Close

Exit Sub

UpRel_Err:
If Err = 3021 Then
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "JOB HISTORY", "HRJOB/PERF", "INSERT")
Call RollBack '26July99 js

End Sub

Private Sub DecSetup()
If glbCompDecHR = 3 Then
    medSalary.Format = "#,##0.000;(#,##0.000)"
End If
If glbCompDecHR = 4 Then
    medSalary.Format = "#,##0.0000;(#,##0.0000)"
End If
End Sub

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function

Private Sub updFollow(EEID&, xType)
Dim SQLQ As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err
If xType = "R" Then
    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED=1 ,EF_EMPNBR=" & EEID&
    SQLQ = SQLQ & " WHERE EF_EMPNBR=" & glbTran_ID
    SQLQ = SQLQ & " AND EF_FREAS='TRAN' AND EF_FDATE=" & Date_SQL(rsORG("TL_NEWDIVEDATE"))
    gdbAdoIhr001.Execute SQLQ
End If
If xType = "S" Then
    If IsDate(dlpDate(2)) Then
        rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = EEID&
        rsTB("EF_FDATE") = CVDate(dlpDate(2))
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(EEID&, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "SREV"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        MsgBox "A Follow Up Record was created!"
    End If
End If
Exit Sub

CrFollow_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Function modReinAudit(EID&)
Dim X%, DtTm As Variant

Screen.MousePointer = HOURGLASS

modReinAudit = False
'rsORG("AU_TODIV") = clpDIV & ""
'rsORG("AU_TIDIV") = clpDIV & ""
'rsORG("AU_TIREASON_TABL") = "SDJC"
'rsORG("AU_TIREASON") = clpCode(1)
'rsORG("AU_TIDATE") = dlpDate(0)
'rsORG("AU_TIEMPNBR") = EID&
rsORG("TL_TCOMPLETE") = "Y"
rsORG.Update
'rsORG.Close

modReinAudit = True
Screen.MousePointer = DEFAULT
Exit Function

Err_Msg:
Screen.MousePointer = DEFAULT
MsgBox "Problem Creating Audit record - Termination Aborted"

End Function


Function CheckSINSSNGen(xSINSNN, TypeFlag)
Dim RsSIN As New ADODB.Recordset
Dim SQLQ
If Not glbLinamar Then If xSINSNN = "999999999" Then Exit Function
    CheckSINSSNGen = False
    SQLQ = "SELECT ED_EMPNBR,ED_SIN,ED_SSN FROM HREMP "
    If TypeFlag = "SIN" Then
        SQLQ = SQLQ & "WHERE ED_SIN = '" & xSINSNN & "' "
    Else
        SQLQ = SQLQ & "WHERE ED_SSN = '" & xSINSNN & "' "
    End If

    RsSIN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsSIN.EOF Then
        CheckSINSSNGen = True
    End If
    RsSIN.Close
    
End Function



Private Sub set_CountrySIN(Optional wDIV)
Dim SQLQ, rsTERM As New ADODB.Recordset
SQLQ = "SELECT ED_EMPNBR,ED_COUNTRY,ED_DIV,ED_PROV,ED_SIN FROM TERM_HREMP WHERE TERM_SEQ=" & fglbTERM_Seq
rsTERM.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsTERM.EOF Then
    Screen.MousePointer = DEFAULT
    MsgBox lblEEName & " has already been rehired. You can not transfer this employee."
    rsTERM.Close
    Call UPDMOD
    xUpdateable = False
    MDIMain.panHelp(0).Caption = ""
Else
    Dim oCountry, oProv
    Dim wCountry, wProv
    If IsMissing(wDIV) Then wDIV = rsTERM("ED_DIV")
    oCountry = rsTERM("ED_COUNTRY")
    oProv = rsTERM("ED_PROV")
    Select Case wDIV
    Case "501", "505"
        wCountry = "MEXICO"
        wProv = "CO"
    Case "502"
        wCountry = "MEXICO"
        wProv = "DU"
    Case "117"
        wCountry = "U.S.A."
        wProv = "IA"
    Case "118"
        wCountry = "U.S.A."
        wProv = "IL"
    Case "120"
        wCountry = "U.S.A."
        wProv = "IA"
    Case "345", "310", "370"
        wCountry = "U.S.A."
        wProv = oProv
    Case "360"
        wCountry = "U.S.A."
        wProv = "KY"
    Case "430"
        wCountry = "GERMANY"
        wProv = oProv
    Case Else
        wCountry = "CANADA"
        wProv = "ON"
    End Select
    '''Ticket #22819 Franks 11/23/2012 - begin
    '''users don't want to show the message also no SIN enter here
    ''If oCountry <> wCountry Then
    ''    lblCountry = "Country will changed from """ & oCountry & """ to """ & wCountry & """"
    ''    lblCountry.Visible = True
    ''Else
    ''    lblCountry = ""
    ''    lblCountry.Visible = False
    ''End If
    ''If oProv <> wProv Then
    ''    lblPROV = "Province will changed from """ & oProv & """ to """ & wProv & """"
    ''    lblPROV.Visible = True
    ''Else
    ''    lblPROV = ""
    ''    lblPROV.Visible = False
    ''End If
    ''If gSec_Show_SIN_SSN Then
    ''    If wCountry = "CANADA" Then
    ''        If Not IsNull(rsTERM("ED_SIN")) Then
    ''            If Not SIN_chk(rsTERM("ED_SIN")) Then
    ''                lblSIN.Caption = "Enter S.I.N."
    ''                lblSIN.Visible = True
    ''                medSIN.Tag = "11-Social Insurance Number"   '
    ''                medSIN.Mask = "###-###-###"
    ''                medSIN.Visible = True
    ''            End If
    ''        End If
    ''    End If
    ''End If
    '''Ticket #22819 Franks 11/23/2012 - end
    MDIMain.panHelp(0).Caption = "Complete the screen"
End If

End Sub
Private Sub Set_PositionSalaWFCHRSoft(rsHRSoft As ADODB.Recordset) 'Ticket #24184 Franks 10/28/2013
Dim EMP_Snap As New ADODB.Recordset
Dim HRSH_Snap As New ADODB.Recordset
Dim HRJH_Snap As New ADODB.Recordset
Dim SQLQ, xDATE, xLinDate, NewDate, dtY1%, dtY2%, xSalary, xSalCD
SQLQ = "Select ED_EMPNBR,ED_DOH,ED_SENDTE,ED_LTHIRE,ED_DEPTNO,ED_GLNO,ED_LOC,ED_SECTION,ED_REGION,ED_LTHIRE,ED_BENEFIT_GROUP,ED_SALDIST,ED_SUPCODE,ED_OMERS,ED_VADIM1,ED_USRDAT1 FROM HREMP WHERE ED_EMPNBR = " & glbTran_ID
EMP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not (EMP_Snap.BOF And EMP_Snap.EOF) Then
    xLinDate = EMP_Snap("ED_DOH")
    If Not glbLinamar Then
        If Not IsNull(EMP_Snap("ED_DOH")) Then dlpDate(3) = EMP_Snap("ED_DOH")
        If Not IsNull(EMP_Snap("ED_SENDTE")) Then dlpDate(4) = EMP_Snap("ED_SENDTE")
        If Not IsNull(EMP_Snap("ED_LTHIRE")) Then dlpDate(5) = EMP_Snap("ED_LTHIRE")
    End If
    
    'position
    If Not IsNull(rsHRSoft("SF_POSITIONCODE")) Then clpJob.Text = rsHRSoft("SF_POSITIONCODE")
    
    If IsDate(xLinDate) Then
        SQLQ = "Select * from HR_SALARY_HISTORY WHERE SH_EMPNBR = " & glbTran_ID & " AND SH_CURRENT <>0"
        HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xDATE = 0: xSalary = "": xSalCD = ""
        If Not (HRSH_Snap.BOF And HRSH_Snap.EOF) Then
            xDATE = HRSH_Snap("SH_EDATE")
            If Not IsNull(rsHRSoft("SF_SALARY")) Then xSalary = rsHRSoft("SF_SALARY") Else xSalary = HRSH_Snap("SH_SALARY")
            xSalCD = HRSH_Snap("SH_SALCD")
            If Not IsNull(rsHRSoft("SF_SALARYFREQUENCY")) Then
                If rsHRSoft("SF_SALARYFREQUENCY") = "Annum" Then xSalCD = "A"
                If rsHRSoft("SF_SALARYFREQUENCY") = "Hour" Then xSalCD = "H"
                If rsHRSoft("SF_SALARYFREQUENCY") = "Monthly" Then xSalCD = "M"
                If rsHRSoft("SF_SALARYFREQUENCY") = "Daily" Then xSalCD = "D"
            End If

            If glbWFC Then
                If Not IsNull(HRSH_Snap("SH_FISCALYEAR")) Then
                    txtFiscalYear.Text = HRSH_Snap("SH_FISCALYEAR")
                End If
                If Not IsNull(HRSH_Snap("SH_MARKETLINE")) Then
                    txtMarketLine.Text = HRSH_Snap("SH_MARKETLINE")
                End If
                If Not IsNull(HRSH_Snap("SH_PAYP")) Then 'Ticket #15818
                    clpCode(9).Text = HRSH_Snap("SH_PAYP")
                End If
                Call Set_MarketLine_List
            End If
        End If
        HRSH_Snap.Close
        If IsDate(xDATE) Then
            dtY1% = DateDiff("yyyy", CVDate(xLinDate), CVDate(xDATE))
            NewDate = DateAdd("yyyy", (dtY1% + 1), CVDate(xLinDate))
        Else
            NewDate = DateAdd("m", 3, CVDate(xLinDate))
        End If
        If Not glbLinamar Then  'Ticket #18364  - They do not want to display the salary
            medSalary = xSalary
        End If
        comPayPer.ListIndex = IIf(xSalCD = "", -1, IIf(xSalCD = "A", 0, IIf(xSalCD = "H", 1, IIf(xSalCD = "M", 2, 3))))
        lblSalCode = xSalCD

    End If
End If
EMP_Snap.Close

End Sub

Private Sub Set_PositionSalary()
Dim EMP_Snap As New ADODB.Recordset
Dim HRSH_Snap As New ADODB.Recordset
Dim HRJH_Snap As New ADODB.Recordset
Dim SQLQ, xDATE, xLinDate, NewDate, dtY1%, dtY2%, xSalary, xSalCD

If glbDivTranInPlant = "Y" Then 'Ticket #25221 Franks 03/17/2014
    Call WFC_DivTranEmpPosSalary
    Exit Sub
End If

SQLQ = "Select ED_EMPNBR,ED_DOH,ED_SENDTE,ED_LTHIRE,ED_DEPTNO,ED_GLNO,ED_LOC,ED_SECTION,ED_REGION,ED_LTHIRE,ED_BENEFIT_GROUP,ED_SALDIST,ED_SUPCODE,ED_OMERS,ED_VADIM1,ED_USRDAT1 FROM TERM_HREMP WHERE TERM_SEQ=" & fglbTERM_Seq
EMP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not (EMP_Snap.BOF And EMP_Snap.EOF) Then
    xLinDate = EMP_Snap("ED_DOH")
    If Not glbLinamar Then
        If Not IsNull(EMP_Snap("ED_DOH")) Then dlpDate(3) = EMP_Snap("ED_DOH")
        If Not IsNull(EMP_Snap("ED_SENDTE")) Then dlpDate(4) = EMP_Snap("ED_SENDTE")
        If Not IsNull(EMP_Snap("ED_LTHIRE")) Then dlpDate(5) = EMP_Snap("ED_LTHIRE")
    End If
    If glbSamuel Then 'Ticket #21791 Franks 04/02/2012
        If Not IsNull(EMP_Snap("ED_DEPTNO")) Then clpDept.Text = EMP_Snap("ED_DEPTNO")
        If Not IsNull(EMP_Snap("ED_GLNO")) Then clpGLNum.Text = EMP_Snap("ED_GLNO")
        If Not IsNull(EMP_Snap("ED_LOC")) Then clpCode(7).Text = EMP_Snap("ED_LOC")
        If Not IsNull(EMP_Snap("ED_SECTION")) Then clpCode(0).Text = EMP_Snap("ED_SECTION")
        If Not IsNull(EMP_Snap("ED_REGION")) Then clpCode(6).Text = EMP_Snap("ED_REGION")
        If Not IsNull(EMP_Snap("ED_LTHIRE")) Then dlpDate(5).Text = EMP_Snap("ED_LTHIRE")
        'Ticket #22260 Franks 07/27/2012 - begin
        If Not IsNull(EMP_Snap("ED_BENEFIT_GROUP")) Then clpBGroup.Text = EMP_Snap("ED_BENEFIT_GROUP")
        If Not IsNull(EMP_Snap("ED_SALDIST")) Then clpSalDist.Text = EMP_Snap("ED_SALDIST")
        If Not IsNull(EMP_Snap("ED_SUPCODE")) Then clpCode(11).Text = EMP_Snap("ED_SUPCODE")
        If Not IsNull(EMP_Snap("ED_OMERS")) Then dlpDate(7).Text = EMP_Snap("ED_OMERS")
        If Not IsNull(EMP_Snap("ED_VADIM1")) Then clpVadim1.Text = EMP_Snap("ED_VADIM1")
        If Not IsNull(EMP_Snap("ED_USRDAT1")) Then dlpDate(7).Text = EMP_Snap("ED_USRDAT1")
        'Ticket #22260 Franks 07/27/2012 - end
    End If
    If IsDate(xLinDate) Then
        SQLQ = "Select * from TERM_SALARY_HISTORY WHERE TERM_SEQ=" & fglbTERM_Seq & " AND SH_CURRENT <>0"
        HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xDATE = 0: xSalary = "": xSalCD = ""
        If Not (HRSH_Snap.BOF And HRSH_Snap.EOF) Then
            xDATE = HRSH_Snap("SH_EDATE")
            xSalary = HRSH_Snap("SH_SALARY")
            If IsNull(HRSH_Snap("SH_SALCD")) Then 'Ticket #29406 Frank 11/01/2016
                xSalCD = ""
            Else
                xSalCD = HRSH_Snap("SH_SALCD")
            End If
            
            If glbWFC Then
                If Not IsNull(HRSH_Snap("SH_FISCALYEAR")) Then
                    txtFiscalYear.Text = HRSH_Snap("SH_FISCALYEAR")
                End If
                If Not IsNull(HRSH_Snap("SH_MARKETLINE")) Then
                    txtMarketLine.Text = HRSH_Snap("SH_MARKETLINE")
                End If
                If Not IsNull(HRSH_Snap("SH_PAYP")) Then 'Ticket #15818
                    clpCode(9).Text = HRSH_Snap("SH_PAYP")
                End If
                Call Set_MarketLine_List
            End If
            If glbSamuel Then 'Ticket #21791 Franks 04/02/2012
                If Not IsNull(HRSH_Snap("SH_EDATE")) Then dlpDate(2).Text = HRSH_Snap("SH_EDATE")
                If Not IsNull(HRSH_Snap("SH_SREAS1")) Then clpCode(8).Text = HRSH_Snap("SH_SREAS1")
                If Not IsNull(HRSH_Snap("SH_PAYP")) Then clpCode(9).Text = HRSH_Snap("SH_PAYP")
                If Not IsNull(HRSH_Snap("SH_GRID")) Then clpGrid.Text = HRSH_Snap("SH_GRID")
            End If
        End If
        HRSH_Snap.Close
        If IsDate(xDATE) Then
            dtY1% = DateDiff("yyyy", CVDate(xLinDate), CVDate(xDATE))
            NewDate = DateAdd("yyyy", (dtY1% + 1), CVDate(xLinDate))
        Else
            NewDate = DateAdd("m", 3, CVDate(xLinDate))
        End If
        If Not glbLinamar Then  'Ticket #18364  - They do not want to display the salary
            medSalary = xSalary
        End If
        comPayPer.ListIndex = IIf(xSalCD = "", -1, IIf(xSalCD = "A", 0, IIf(xSalCD = "H", 1, IIf(xSalCD = "M", 2, 3))))
        lblSalCode = xSalCD
        If glbLinamar Then
            dlpDate(2) = NewDate
        End If
        
        If Not glbLinamar Then  'Ticket #18364  - They do not want to display the Job
            SQLQ = "Select * from TERM_JOB_HISTORY WHERE TERM_SEQ=" & fglbTERM_Seq & " AND JH_CURRENT <>0"
            HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not HRJH_Snap.EOF Then
                If glbLinamar Then
                    clpJob = HRJH_Snap("JH_JOB")
                End If
                If glbSamuel Then 'Ticket #21791 Franks 04/02/2012
                    clpJob = HRJH_Snap("JH_JOB")
                    If Not IsNull(HRJH_Snap("JH_DHRS")) Then medHours(0).Text = HRJH_Snap("JH_DHRS")
                    If Not IsNull(HRJH_Snap("JH_WHRS")) Then medHours(1).Text = HRJH_Snap("JH_WHRS")
                    If Not IsNull(HRJH_Snap("JH_PHRS")) Then medHours(2).Text = HRJH_Snap("JH_PHRS")
                    If Not IsNull(HRJH_Snap("JH_REPTAU")) Then elpReptAuthShow(0).Text = HRJH_Snap("JH_REPTAU")
                    If Not IsNull(HRJH_Snap("JH_REPTAU2")) Then elpReptAuthShow(1).Text = HRJH_Snap("JH_REPTAU2")
                    If Not IsNull(HRJH_Snap("JH_REPTAU3")) Then elpReptAuthShow(2).Text = HRJH_Snap("JH_REPTAU3")
                    If Not IsNull(HRJH_Snap("JH_REPTAU4")) Then elpReptAuthShow(3).Text = HRJH_Snap("JH_REPTAU4")
                    If HRJH_Snap("JH_PROFIT_SHARING") Then chkProSha.Value = 1
                End If
            End If
        End If
    End If
End If
EMP_Snap.Close

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
Updateble = xUpdateable
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
    Call UPDMOD
End If

End Sub

Sub Set_MarketLine_List()
Dim rsWFC As New ADODB.Recordset
Dim X%, I%
Dim xItemAdd
Dim SQLQ

If Not glbWFC Then Exit Sub

SQLQ = "select MarketLine from WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & fglbBAND & "'"
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
Sub cmbMarketLine_GotFocus()   'Jaddy 8/9/99
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub cmbMarketLine_LostFocus()
    txtMarketLine = cmbMarketLine
End Sub

Private Sub txtEmpID_LostFocus()
If glbSamuel Then
    Call EmpNoExist(getEmpnbr(txtEmpID.Text))
End If
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
Call WFC_Band 'Ticket #21677 Franks 03/07/2012
Call Set_MarketLine_List

End Sub

Private Sub txtLabCode_Change()
lblLabCodeDesc.Caption = getLabCodeDesc(txtLabCode.Text)
If Len(txtLabCode.Text) > 0 Then
    lblLabCodeDesc.Visible = True
Else
    lblLabCodeDesc.Visible = False
End If
End Sub
Private Function getLabCodeDesc(xCode)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = "Unassigned"
    If Not IsNull(xCode) Then
        SQLQ = "SELECT TB_NAME, TB_KEY, TB_DESC FROM HRTABL WHERE TB_NAME = 'SDLB' AND TB_KEY = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("TB_DESC")
        End If
        rsDiv.Close
    End If
    getLabCodeDesc = xRetVal
End Function
Private Sub txtLabCode_DblClick()
    Call Get_Code_Normal("SDLB", "Labour Code", "")
    If Len(glbCode) > 0 Then
        txtLabCode.Text = glbCode
    End If
End Sub

Private Sub txtLabCode_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtMarketLine_Change() 'Jaddy 8/9/99
  'cmbMarketLine.Clear
  'MarketLine_AddItem Me
  'setMarketLine Me
  Call SalMarketLineDesc

End Sub

Function getProductLineCodeforLinamar(xOrgCode)
    Dim rsTABL As New ADODB.Recordset
    Dim xNewCode
    xNewCode = xOrgCode
    rsTABL.Open "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDRG' AND TB_KEY='" & xOrgCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If rsTABL.EOF Or rsTABL.BOF Then
        xNewCode = "ALL" & Mid(xOrgCode, 4)
    End If
    getProductLineCodeforLinamar = xNewCode
End Function

Private Sub comEmpType_Click()
'comEmpType.Sorted = True before
If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2384W" Then 'St. John's Rehab Hospital 'Ticket #14752
    'St. Marys
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
comEmpType.AddItem "6 - Contract Salary"
comEmpType.AddItem "7 - Contract Hourly"    '23June99 js
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    'comEmpType.AddItem "8 - 80% Full Time Salary"
    comEmpType.AddItem "8 - 80% @7.5hrs"   'Ticket #14995
Else
    comEmpType.AddItem "8 - Salary Pensioners"
End If
'comEmpType.AddItem "9 - Salary Elected officials"
'Added by Bryan 12/08/05 for granite club
If glbCompSerial = "S/N - 2241W" Then
    comEmpType.AddItem "9 - Commissioned Employees"
ElseIf glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    comEmpType.AddItem "9 - Former Air Liquide Emp."
Else
    comEmpType.AddItem "9 - Salary Elected officials"
End If
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14736
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

End Sub

Private Sub WFC_NGS_Trans(xEmpNo, xType) '#19266
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xTranInDate
    
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    ''this function was done in Transfer Out
    '''Ticket #24620 Franks 12/03/2013
    '''if you were transferring from NGS to Canada or any other non US plant:"   NGS End Date = Transfer Out Date
    ''If IsDate(xWFC_OldNGSStart) Then
    ''    If Not IsWFCNGSDiv(clpDiv.Text) Then
    ''        Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", DateAdd("d", -1, CVDate(dlpDate(1).Text)), "Y")
    ''        Exit Sub
    ''    End If
    ''End If
    
    glbWFCNGSSubGroup = ""
    'Ticket #23247 Franks 07/23/2013 - add ED_EMP
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2, ED_EMP FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        ''get the glbWFCNGSSubGroup and glbWFCPayGroup
        'glbWFCNGSSubGroup = getNGSSubGrpFromMatrix(glbEmpDiv, glbUNION)
        'Ticket #23247 Franks 07/23/2013
        glbWFCNGSSubGroup = getNGSFieldFromMatrix(glbEmpDiv, glbUNION, rsEmpee("ED_EMP"), "NG_SUB_GROUP")
        
        If Len(glbWFCNGSSubGroup) > 0 Then
            If clpCode(12).Text = "COOP" Or clpCode(12).Text = "STUD" Then
                'Ticket #25352 Franks 04/16/2014 - If Employment Status = COOP or STUD, no NGS Sub Code
                rsEmpee("ED_VADIM1") = Null
            Else
                rsEmpee("ED_VADIM1") = glbWFCNGSSubGroup
            End If
            If Len(clpVadim2.Text) = 0 Then
                'glbWFCPayGroup = getNGSPayGrpFromMatrix(glbEmpDiv, glbUNION)
                'Ticket #23247 Franks 07/23/2013
                glbWFCPayGroup = getNGSFieldFromMatrix(glbEmpDiv, glbUNION, rsEmpee("ED_EMP"), "NG_PAY_GROUP")
                If Len(glbWFCPayGroup) > 0 Then
                    rsEmpee("ED_VADIM2") = glbWFCPayGroup
                End If
            Else
                glbWFCPayGroup = clpVadim2.Text
            End If
            rsEmpee.Update
        End If
    End If
    rsEmpee.Close
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    If clpCode(12).Text = "COOP" Or clpCode(12).Text = "STUD" Then
        'Ticket #25352 Franks 04/16/2014 - If Employment Status = COOP or STUD, no NGS Sub Code and Start Date
        Exit Sub
    End If
            
    xLDate = dlpDate(1).Text 'Date
    xTranInDate = dlpDate(1).Text
    
    'xNGSStart = dlpDate(1).Text
    xNGSStart = dlpDate(6).Text
    If Not IsDate(xNGSStart) Then Exit Sub
    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmpOther.EOF Then
        rsEmpOther.AddNew
        rsEmpOther("ER_COMPNO") = "001"
        rsEmpOther("ER_EMPNBR") = xEmpNo
    End If
    
    rsEmpOther("ER_OTHERDATE1") = CVDate(xNGSStart)
    rsEmpOther("ER_LDATE") = Date
    rsEmpOther("ER_LTIME") = Time$
    rsEmpOther("ER_LUSER") = glbUserID
    rsEmpOther.Update
    rsEmpOther.Close
    'No NGS Effective Date, skip
    If Len(xNGSStart) = 0 Then Exit Sub

    If glbUNION = "NONE" Or glbUNION = "EXEC" Then
        xSalHly = "Y"
    Else
        xSalHly = "N"
    End If

    Call NGSAuditAdd(xEmpNo, "M", "Transfer In", "Transfer In Date", "", CVDate(xTranInDate), xLDate)
    Call NGSAuditAdd(xEmpNo, "M", "Transfer In", "New Division", "", glbEmpDiv, xLDate)
    Call NGSAuditAdd(xEmpNo, "M", "Transfer In", lStr("Vadim Field 1"), "", glbWFCNGSSubGroup, xLDate)
    'Ticket #19955 Franks 03/07/2011
    If Not dlpDate(6).Visible Then
        Call NGSAuditAdd(xEmpNo, "M", "Transfer In", lStr("Other Date 1"), "", CVDate(xNGSStart), xLDate)
    Else
        Call NGSAuditAdd(xEmpNo, "M", "Transfer In", lStr("Other Date 1"), xWFC_OldNGSStart, CVDate(xWFC_NewNGSStart), xLDate)
    End If

End Sub

Private Sub LinamarScreenSetup() 'Ticket #29414 Franks 11/03/2016
    clpCode(1) = "TRAN"
    txtEmpID = Left(glbTran_ID, Len(glbTran_ID & "") - 3)
    clpDIV.MaxLength = 3
    clpCode(2).MaxLength = 8
    clpHOME(1).MaxLength = 12
    clpHOME(2).MaxLength = 12
    clpHOME(3).MaxLength = 12
    clpHOME(4).MaxLength = 5
    txtPayrollID.MaxLength = 15
    lbltitle(7).FontBold = True ' for product line
    
    'Ticket #29414 Franks 11/03/2016 - new begin
    lblReptAuth(0).Top = 4560 'Rept. Authority 1
    lblReptAuth(0).Left = 5820
    lblReptAuth(0).FontBold = True
    lblReptAuth(0).Visible = True
    elpReptAuthShow(0).Top = 4560
    elpReptAuthShow(0).Left = 7500
    elpReptAuthShow(0).Visible = True

    lblShift.Top = 4920 'Shift
    txtShift.Top = 4920
    clpCode(13).Visible = True
    clpCode(13).Top = 4920
    clpCode(13).Left = 7500
    
    lbltitle(14).Top = 5280 'Salary
    medSalary.Top = 5280

    lbltitle(22).Top = 5610 'Per
    comPayPer.Top = 5610 '5610
    
    lbltitle(5).Top = 5940 'Next Review Date
    dlpDate(2).Top = 5940

    lbltitle(27).Top = 6270 ' Labour Code
    clpCode(4).Top = 6270
    'Ticket #29946 Franks 03/15/2017 - begin
    clpCode(4).Visible = False
    frmLinLabourCode.Top = clpCode(4).Top
    frmLinLabourCode.Left = clpCode(4).Left
    frmLinLabourCode.Visible = True
    frmLinLabourCode.BorderStyle = 0
    'Ticket #29946 Franks 03/15/2017 - end
    'Ticket #29414 Franks 11/03/2016 - new end
    
    'Ticket #29759 Franks 02/21/2017 - begin
    'for Payroll ID
    txtPayrollID.Enabled = False
    txtPayrollID.Text = getTermEmpPayID
    cmdEditPayID.Visible = True
    
    xPayIDEnable = False
    'Ticket #29759 Franks 02/21/2017 - end
End Sub

Private Sub WFCScreenSetup() 'Ticket #22448
Dim X%, SQLQ
Dim xOrga1
Dim xVal
    xVal = 360
    
    clpCode(1).Enabled = True
    'clpCode(1) = "TRNI"
    
    'Ticket #24695 Franks 11/28/2013
    'txtEmpID = Right(glbTran_ID, Len(glbTran_ID & "") - 4)
    txtEmpID = glbTran_ID
    lblEmpNo.Visible = False
    
    lblShift.Visible = False
    txtShift.Visible = False
    For X% = 7 To 13
        If X% <> 10 Then
        lbltitle(X%).Visible = False
        End If
    Next X%
    lbltitle(27).Visible = False
    clpCode(4).Visible = False
    clpCode(2).Visible = False
    clpHOME(1).Visible = False
    clpHOME(2).Visible = False
    'clpGLNum.Visible = False
    clpHOME(3).Visible = False
    clpHOME(4).Visible = False
    clpCode(3).Visible = False
    lblSection.Visible = True
    clpCode(0).Visible = True
    lbltitle(23).Visible = True
    lbltitle(24).Visible = True
    lbltitle(25).Visible = True
    clpCode(5).Visible = True
    clpCode(6).Visible = True
    clpCode(7).Visible = True
    lblEEStatus.Visible = True
    clpCode(8).Visible = True
    lblOHire.Visible = True
    dlpDate(3).Visible = True
    lblSen.Visible = True
    dlpDate(4).Visible = True
    lblLHire.Visible = True
    dlpDate(5).Visible = True
    lbltitle(28).Visible = True 'Ticket #15818
    clpCode(9).Visible = True   'Ticket #15818
    lblSection.Top = 3120 '3360
    clpCode(0).Top = 3120
    lbltitle(6).Top = 3480 '3720
    clpDept.Top = 3480
    lbltitle(10).Top = 3840 '4080
    clpGLNum.Top = 3840
    lbltitle(23).Top = 4200 '4440
    clpCode(7).Top = 4200 '4440 'Location
    lbltitle(25).Top = 4560 '4800
    clpCode(5).Top = 4560
    lbltitle(24).Top = 4920 '5160
    clpCode(6).Top = 4920 'Region - Home Shift
    lblOHire.Top = 5280 '5520 'Original Hire Date - Operation
    dlpDate(3).Top = 5280
    lblSen.Top = 5640 '5880
    dlpDate(4).Top = 5640 'Seniority Date
    lblLHire.Top = 6000 '6240
    dlpDate(5).Top = 6000 'Last Hire Date
    
    lblSection.Caption = lStr(lblSection.Caption)
    lbltitle(5).FontBold = False
    clpJob.Enabled = True
    lblEEStatus.Top = 5610 '5850
    clpCode(8).Top = 5610 '5850 'Reason for Change
    If glbWFC Then
        clpCode(8).Enabled = False 'For WFC, user can't chang it
        clpCode(8).Text = "TRNI"
        txtEmpID.Enabled = False 'Ticket #15895
        lbltitle(2).Caption = "New Payroll ID"
    End If
    lbltitle(26).FontBold = True
    If glbWFC Then
        lbltitle(28).Top = 6645 '#15818
        clpCode(9).Top = 6645 '#15818
        'If rsORG("ED_ORG") = "NONE" Or rsORG("ED_ORG") = "EXEC" Then
        Call WFC_PosSalScreen ''Ticket #24936 Franks 02/05/2014
        Call WFC_UnionScreen(rsORG("ED_ORG"))
        
        'Populate some fields using Division Master - Begin Ticket #15514
        xOrga1 = ""
        If Len(clpDIV.Text) > 0 Then
            Dim rsTTemp As New ADODB.Recordset
            SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDIV.Text & "' "
            rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTTemp.EOF Then
                If Not IsNull(rsTTemp("DV_SECTION")) Then clpCode(0).Text = rsTTemp("DV_SECTION")
                If Not IsNull(rsTTemp("DV_LOC")) Then clpCode(7).Text = rsTTemp("DV_LOC")
                If Not IsNull(rsTTemp("DV_REGION")) Then clpCode(6).Text = rsTTemp("DV_REGION")
                If Not IsNull(rsTTemp("DV_ADMINBY")) Then clpCode(5).Text = rsTTemp("DV_ADMINBY")
                If Not IsNull(rsTTemp("DV_ORGT1")) Then xOrga1 = rsTTemp("DV_ORGT1")
            End If
            rsTTemp.Close
        End If
        'Populate some fields using Division Master - End
        
        'Ticket #16748
        lblVadim2.Visible = True
        clpVadim2.Visible = True
        lblVadim2.Top = clpCode(9).Top + 360
        lblVadim2.Caption = lStr("Vadim Field 2")

        'Ticket #30376 Franks 07/17/2017 - begin
        lbltitle(31).Visible = True
        clpCode(14).Visible = True
        'lbltitle(31).Top = lblVadim2.Top + 360 'Ticket #30446 Franks 08/09/2017
        lbltitle(31).Left = lblVadim2.Left
        lbltitle(31).Caption = lStr("Organization 1")
        'clpCode(14).Top = clpVadim2.Top + 360 'Ticket #30446 Franks 08/09/2017
        clpCode(14).Left = clpVadim2.Left
        If clpCode(0).Text = "MISS" Or clpCode(0).Text = "TROY" Then
            lbltitle(31).Enabled = True: clpCode(14).Enabled = True
            clpCode(14).Text = xOrga1
        Else
            lbltitle(31).Enabled = False: clpCode(14).Enabled = False
        End If
        lblWFCNote5.Visible = True
        'Ticket #30376 Franks 07/17/2017 - end
        
        ''Ticket #23247 Franks 09/13/2013
        'lblVadim11.Left = lblVadim2.Left
        'lblVadim11.Top = lblVadim2.Top + 360
        'clpVadim1.Left = clpVadim2.Left
        'clpVadim1.Top = clpVadim2.Top + 360
        'lblVadim11.Visible = True
        'clpVadim1.Visible = True
        'lblVadim11.Caption = lStr("Vadim Field 1")
        
        'Ticket #16395 - 09/16/2009 Pension II - Begin
        lblEEType.Top = 6330
        comEmpType.Top = 6330
        comEmpType.Left = medSIN.Left - 10
        lblEEType.Visible = True '
        comEmpType.Visible = True
        
        'Ticket #23247 Franks 07/23/2013 - add Emp Status - begin
        lblEEStatus2.Top = 6675
        lblEEStatus2.Left = lblLHire.Left
        clpCode(12).Top = 6675
        clpCode(12).Left = clpCode(7).Left
        lblEEStatus2.Visible = True
        clpCode(12).Visible = True
        lblEEStatus2.Caption = lStr("Employment Status")
        'Ticket #23247 Franks 07/23/2013 - add Emp Status - end
        
        lblUnion.Left = lblLHire.Left
        lblEEType.Left = lblLHire.Left
        lblUnion.Top = 6675 + 320
        clpCode(10).Top = 6675 + 320
        clpCode(10).Left = clpCode(7).Left
        lblUnion.Visible = True
        clpCode(10).Visible = True
        
        'Ticket #18654 - begin
        lblBen.Left = lblLHire.Left
        lblBen.Top = 6995 + 320
        lblBen.Visible = True
        clpBGroup.Left = clpCode(7).Left
        clpBGroup.Top = 6995 + 320
        clpBGroup.Visible = True
        If GetCountryFromDiv(clpDIV.Text) = "CANADA" Then
            lblBen.FontBold = True
            lblUserNum1.FontBold = True
            lblUserText2.FontBold = True
        End If
        'Ticket #18654 - end
        
        'Ticket #22448 - begin
        lblUserNum1.Top = 7320 + 320
        txtUserNum1.Top = 7320 + 320
        lblUserText2.Top = 7630 + 320
        comUserText2.Top = 7630 + 320
        lblUserText2.Left = lblUserNum1.Left
        txtUserNum1.Left = medSIN.Left - 10
        comUserText2.Left = medSIN.Left - 10
        comUserText2.Width = 3000
        lblUserNum1.Visible = True
        txtUserNum1.Visible = True
        lblUserText2.Visible = True
        comUserText2.Visible = True
        lblUserNum1.Caption = lStr("User Number 1")
        lblUserText2.Caption = lStr("User Text 2")
        'Ticket #22448 - end
        
        Call ComEType
        lblEEType.Caption = lStr("Employment Type")
        lblUnion.Caption = lStr("Union")
        OldUnion = ""
        'Ticket #25088 Franks 02/18/2014 - begin
        OldSection = ""
        OldRegion = ""
        OldAdminBy = ""
        OldLoc = ""
        If Not IsNull(rsORG("ED_SECTION")) Then OldSection = rsORG("ED_SECTION")
        If Not IsNull(rsORG("ED_REGION")) Then OldRegion = rsORG("ED_REGION")
        If Not IsNull(rsORG("ED_ADMINBY")) Then OldAdminBy = rsORG("ED_ADMINBY")
        If Not IsNull(rsORG("ED_LOC")) Then OldLoc = rsORG("ED_LOC")
        'Ticket #25088 Franks 02/18/2014 - end
        
        If Not IsNull(rsORG("ED_ORG")) Then
            clpCode(10).Text = rsORG("ED_ORG")
            OldUnion = rsORG("ED_ORG")
        End If
        If Not IsNull(rsORG("ED_DEPTNO")) Then OldDept = rsORG("ED_DEPTNO") Else OldDept = ""
        If glbCandidate > 0 Then 'Ticket #24184 Franks 10/28/2013
            If Not IsNull(rsORG("ED_DIV")) Then OldDiv = rsORG("ED_DIV") Else OldDiv = ""
        Else
            If Not IsNull(rsORG("TL_OLDDIV")) Then OldDiv = rsORG("TL_OLDDIV") Else OldDiv = ""
        End If
        'Ticket #21677 Franks 03/15/2012 - begin 'rsORG!
        SaveBGroup = ""
        If Not IsNull(rsORG("ED_BENEFIT_GROUP")) Then
            SaveBGroup = rsORG("ED_BENEFIT_GROUP")
        End If
        locSurName = ""
        If Not IsNull(rsORG("ED_SURNAME")) Then
            locSurName = rsORG("ED_SURNAME")
        End If
        locFName = ""
        If Not IsNull(rsORG("ED_FNAME")) Then
            locFName = rsORG("ED_FNAME")
        End If
        'Ticket #21677 Franks 03/15/2012 - end
        
        If Not IsNull(rsORG("ED_EMPTYPE")) Then
            If rsORG("ED_EMPTYPE") = "Y" Then
                comEmpType.ListIndex = 0
            End If
            If rsORG("ED_EMPTYPE") = "N" Then
                comEmpType.ListIndex = 1
            End If
        End If
        Call comEmpType_Click
        'Ticket #16395 - 09/16/2009 Pension II - End
        
        'Ticket #19955 Franks 03/07/2011 - begin
        xWFC_OldNGSStart = ""
        
        'If Not IsNull(rsORG("ED_WORKCOUNTRY")) Then
        If IsWFCNGSDiv(clpDIV.Text) Then 'Ticket #24620 Franks 12/03/2013
            'If rsORG("ED_WORKCOUNTRY") = "U.S.A." Then
                If clpDIV.Text = "1094" Then 'Ticket #24451 Franks 10/15/2013
                    '"   If transferring into GREENSBORO from a NGS plant, don't display the NGS Start Date. The user cannot enter into this field.
                Else
                    lbltitle(29).Caption = lStr("Other Date 1")
                    dlpDate(6).Tag = lStr("Other Date 1")
                    lbltitle(29).Visible = True
                    dlpDate(6).Visible = True
                    lblNGSStart.Visible = True
                    xWFC_OldNGSStart = get_EmpOtherByField("", "ER_OTHERDATE1", glbTran_Seq)
                    
                    If OldDiv = clpDIV.Text And Not (OldUnion = clpCode(10).Text) Then
                        'Ticket #24652 Franks 12/02/2013
                        '"   If the Transfer Out and In Divisions are the same but the Union code is different, the NGS Start Date needs to be Transfer In Date
                        dlpDate(6).Text = dlpDate(1).Text
                    Else
                        If Not IsNull(xWFC_OldNGSStart) Then
                            dlpDate(6).Text = xWFC_OldNGSStart
                        Else
                            xWFC_OldNGSStart = ""
                        End If
                    End If
                    
                End If
            'End If
        End If
        'Ticket #19955 Franks 03/07/2011 - end
        
        '''Ticket #23247 Franks 07/23/2013 - add Emp Status - begin
        ''lblEEStatus2.Top = 7970
        ''lblEEStatus2.Left = lblUserNum1.Left
        ''clpCode(12).Top = 7970
        ''clpCode(12).Left = clpDept.Left
        ''lblEEStatus2.Visible = True
        ''clpCode(12).Visible = True
        ''lblEEStatus2.Caption = lStr("Employment Status")
        '''Ticket #23247 Franks 07/23/2013 - add Emp Status - end
    End If
    '-------------WFC End
    lblEEStatus2.FontBold = True
    lbltitle(10).FontBold = True 'GL
    lblUnion.FontBold = True

    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'For Frank Test with WFC database
        cmdFrankTest.Visible = True
    End If
    
End Sub

Private Sub SamuelScreenSetup()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTERM As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim X%, SQLQ
Dim xIn As Integer

    clpCode(1).Enabled = True
    'clpCode(1) = "TRNI"
    txtEmpID = glbTran_ID ' Right(glbTran_ID, Len(glbTran_ID & "") - 4)
    lblShift.Visible = False
    txtShift.Visible = False
    For X% = 7 To 13
        If X% <> 10 Then
        lbltitle(X%).Visible = False
        End If
    Next X%
    lbltitle(27).Visible = False
    clpCode(4).Visible = False
    clpCode(2).Visible = False
    clpHOME(1).Visible = False
    clpHOME(2).Visible = False
    'clpGLNum.Visible = False
    clpHOME(3).Visible = False
    clpHOME(4).Visible = False
    clpCode(3).Visible = False
    lblSection.Visible = True
    clpCode(0).Visible = True
    lbltitle(23).Visible = True
    lbltitle(24).Visible = True
    lbltitle(25).Visible = True
    clpCode(5).Visible = True
    clpCode(6).Visible = True
    clpCode(7).Visible = True
    lblEEStatus.Visible = True
    clpCode(8).Visible = True
    lblOHire.Visible = True
    dlpDate(3).Visible = True
    lblSen.Visible = True
    dlpDate(4).Visible = True
    lblLHire.Visible = True
    dlpDate(5).Visible = True
    lbltitle(28).Visible = True 'Ticket #15818
    clpCode(9).Visible = True   'Ticket #15818
    lblSection.Top = 4560 ' 3120 '3360
    clpCode(0).Top = 4560 ' 3120
    lbltitle(21).Top = 3480 'Div
    clpDIV.Top = 3480 'Div
    
    lbltitle(10).Top = 3840 '4080
    clpGLNum.Top = 3840
    lbltitle(23).Top = 4200 '4440
    clpCode(7).Top = 4200 '4440 'Location
    lbltitle(25).Top = 1680 'Admin
    clpCode(5).Top = 1680 'Admin
    lblSection.Caption = lStr(lblSection.Caption)
    clpJob.Enabled = True
    lblEEStatus.Top = 5610
    clpCode(8).Top = 5610 'Reason for Change
    
    
    lbltitle(24).Top = 4920 '5160
    clpCode(6).Top = 4920 'Region - Home Shift
    lblOHire.Top = 5640 '5520 'Original Hire Date - Operation
    dlpDate(3).Top = 5640
    lblSen.Top = 6000 '5880
    dlpDate(4).Top = 6000 'Seniority Date
    lblLHire.Top = 6360 '6240
    dlpDate(5).Top = 6360 'Last Hire Date

    'Call WFC_UnionScreen(rsORG("ED_ORG"))
    'Hiding some NONE and EXEC fields
    txtFiscalYear.Visible = False
    cmbMarketLine.Visible = False
    lblFiscalYear.Visible = False
    lblMarketLine.Visible = False
    lblMLine.Visible = False
    lbltitle(28).Top = 5940 '#15818
    clpCode(9).Top = 5940 '#15818
    lblGrid.Top = 6270
    clpGrid.Top = 6270

    lblUnion.Left = lblLHire.Left
    lblEEType.Left = lblLHire.Left
    lblUnion.Top = 5280 '6675
    clpCode(10).Top = 5280 '6675
    clpCode(10).Left = clpCode(7).Left
    lblUnion.Visible = True
    clpCode(10).Visible = True

    'Ticket #21791 Franks 04/09/2012 - begin
    lblReptAuth(0).Left = lbltitle(19).Left 'position cdode
    lblReptAuth(1).Left = lbltitle(19).Left
    lblReptAuth(2).Left = lbltitle(19).Left
    lblReptAuth(3).Left = lbltitle(19).Left
    lbltitle(30).Left = lbltitle(19).Left
    chkProSha.Left = medSalary.Left
    elpReptAuthShow(0).Left = clpJob.Left
    elpReptAuthShow(1).Left = clpJob.Left
    elpReptAuthShow(2).Left = clpJob.Left
    elpReptAuthShow(3).Left = clpJob.Left
    xIn = 330
    For X% = 0 To 3
        lblReptAuth(X%).Top = lblGrid.Top + xIn * (X% + 1)
        elpReptAuthShow(X%).Top = lblGrid.Top + xIn * (X% + 1)
    Next
    lbltitle(30).Top = lblGrid.Top + xIn * 5
    chkProSha.Top = lblGrid.Top + xIn * 5
    
    lblReptAuth(0).Visible = True
    lblReptAuth(1).Visible = True
    lblReptAuth(2).Visible = True
    lblReptAuth(3).Visible = True
    lbltitle(30).Visible = True
    elpReptAuthShow(0).Visible = True
    elpReptAuthShow(1).Visible = True
    elpReptAuthShow(2).Visible = True
    elpReptAuthShow(3).Visible = True
    chkProSha.Visible = True
    'Ticket #21791 Franks 04/09/2012 - end
    
    lblUnion.Caption = lStr("Union")
    If Not rsORG.EOF Then
        If Not IsNull(rsORG("ED_ORG")) Then
            clpCode(10).Text = rsORG("ED_ORG")
        End If
    End If
    If Len(clpCode(5).Text) > 0 Then 'Ticket #21791 Franks 04/09/2012
        If clpCode(5).Text = "5231" Then
            clpCode(10).Text = "EXEC"
        End If
        If clpCode(5).Text = "5230" Or clpCode(5).Text = "5232" Then
            clpCode(10).Text = "NONE"
        End If
    End If
    
    lbltitle(1).Caption = lStr("Division") & " Start Date"
    lbltitle(25).Caption = lStr("Administered By") '"PLANT"
    lblOHire.Caption = lStr("Original Hire")
    lblSen.Caption = lStr("Seniority")
    lblLHire.Caption = lStr("Last Hire")
    lbltitle(5).Caption = "Effective Date"
    
    'Salary Grid Category
    lblGrid.Visible = True
    clpGrid.Visible = True
    lblGrid.Caption = lStr("Grid Category")
    clpGrid.TABLTitle = lStr(lblGrid)
    
    'Ticket #22260 - Franks 07/26/2012 begin #########################
    '---Benefit Group
    lblBen.Caption = lStr("Benefit Group")
    lblBen.Left = lblLHire.Left:    lblBen.Top = 6710:    lblBen.Visible = True
    clpBGroup.Left = clpCode(7).Left: clpBGroup.Top = 6710: clpBGroup.Visible = True
    '---Supervisor Code
    lblSupervisor.Caption = lStr("Supervisor Code")
    lblSupervisor.Left = lblLHire.Left: lblSupervisor.Top = 7040: lblSupervisor.Visible = True
    clpCode(11).Left = clpCode(7).Left: clpCode(11).Top = 7040: clpCode(11).Visible = True
    
    '---Salary Distribution '
    lblSalDist.Caption = lStr("Salary Distribution")
    lblSalDist.Left = lblLHire.Left: lblSalDist.Top = 7370: lblSalDist.Visible = True
    clpSalDist.Left = clpCode(7).Left: clpSalDist.Top = 7370: clpSalDist.Visible = True

    '---OMERS Date
    lblODate.Caption = lStr("OMERS Date")
    lblODate.Left = lblLHire.Left: lblODate.Top = 7700: lblODate.Visible = True
    dlpDate(7).Left = clpCode(7).Left: dlpDate(7).Top = 7700: dlpDate(7).Visible = True
    '---Vadim Field 1
    lblVadim11.Caption = lStr("Vadim Field 1")
    lblVadim11.Left = lblLHire.Left: lblVadim11.Top = 8030: lblVadim11.Visible = True
    clpVadim1.Left = clpCode(7).Left: clpVadim1.Top = 8030: clpVadim1.Visible = True
    '---User Defined
    lblUDay.Caption = lStr("User Defined")
    lblUDay.Left = lblLHire.Left: lblUDay.Top = 8360: lblUDay.Visible = True
    dlpDate(8).Left = clpCode(7).Left: dlpDate(8).Top = 8360: dlpDate(8).Visible = True
    'Ticket #22260 - Franks 07/26/2012 end ##########################
    
    'Font - begin
    lbltitle(21).FontBold = True 'Div
    lbltitle(25).FontBold = True 'Admin
    lbltitle(5).FontBold = True 'section
    lbltitle(23).FontBold = True 'location
    lbltitle(24).FontBold = True 'Region
    lblUnion.FontBold = True
    lbltitle(28).FontBold = True
    'Font - end
    
    'tab order
    dlpDate(0).TabIndex = 0
    dlpDate(1).TabIndex = 1
    clpCode(1).TabIndex = 2
    clpCode(5).TabIndex = 3 'AdminBy
    txtEmpID.TabIndex = 4
    txtPayrollID.TabIndex = 5
    clpDept.TabIndex = 6
    clpDIV.TabIndex = 7
    clpGLNum.TabIndex = 8
    clpCode(7).TabIndex = 9 'Location
    clpCode(0).TabIndex = 10 'section
    clpCode(6).TabIndex = 11 'region
    clpCode(10).TabIndex = 12 'Union
    dlpDate(3).TabIndex = 13
    dlpDate(4).TabIndex = 14
    dlpDate(5).TabIndex = 15
    clpBGroup.TabIndex = 16
    clpCode(11).TabIndex = 17
    clpSalDist.TabIndex = 18
    dlpDate(7).TabIndex = 19
    clpVadim1.TabIndex = 20
    dlpDate(8).TabIndex = 21
    clpJob.TabIndex = 22 '16
    'clpCode(1).SetFocus

End Sub

Private Function getPosBand(xPCODE)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retVal
    retVal = ""
    SQLQ = "SELECT JB_CODE, JB_BAND FROM HRJOB WHERE JB_CODE = '" & xPCODE & "' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("JB_BAND")) Then
            retVal = rsTemp("JB_BAND")
        End If
    End If
    getPosBand = retVal
End Function

Private Sub getPayGroup(xDIV, xUnion)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
    If Len(xUnion) = 0 Then Exit Sub
    If Len(xDIV) = 0 Then Exit Sub
    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
    SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("NG_PAY_GROUP")) Then
            clpVadim2.Text = rsTemp("NG_PAY_GROUP")
        End If
    End If
    rsTemp.Close
End Sub

Private Sub locBeneGroupUpdate(xEmpNo) 'Ticket #21677 Franks 03/15/2012
Dim NewBGroup
Dim Msg
Dim SQLQ
Dim rsEmp As New ADODB.Recordset
Dim rsBenT As New ADODB.Recordset
Dim xTemp

    'This function works for Benefit Group change, so the previous and current
    'benefit group can't be blank
    NewBGroup = clpBGroup
    'If Len(SaveBGroup) = 0 Then Exit Sub
    If Len(NewBGroup) = 0 Then
        'Ticket #21677 Franks 09/16/2013 - begin
        '"1.If the Transfer In Division or Union is not a NGS location (ie: not found in the matrix):
        '"2.Delete the Benefit Group Code
        '"3. Remove the Benefit Group Code from the Company Paid Benefits
        '"   Add a Benefit End Date to equal the Transfer In Date minus 1.
        If Len(SaveBGroup) > 0 Then
            '1.
            SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = NULL WHERE ED_EMPNBR = " & xEmpNo & " "
            gdbAdoIhr001.Execute SQLQ
            '3.
            If IsDate(dlpDate(1).Text) Then
                xTemp = DateAdd("D", -1, CVDate(dlpDate(1).Text))
                'SQLQ = "UPDATE HRBENFT SET BF_CEASEDATE = " & Date_SQL(xTemp) & " WHERE BF_EMPNBR = " & xEmpNo & " "
                'SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                'SQLQ = SQLQ & "AND BF_PCC = 1 "
                'gdbAdoIhr001.Execute SQLQ
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                SQLQ = SQLQ & "AND BF_PCC = 1 "
                'If Not rsBenT.State <> 0 Then rsBenT.Close
                rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsBenT.EOF
                    rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                    rsBenT.Update
                    'update audit
                    Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                    rsBenT.MoveNext
                Loop
            End If
            '2.
                SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE BF_EMPNBR = " & xEmpNo & " "
                'SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                'SQLQ = SQLQ & "AND BF_PCC = 1 "
                gdbAdoIhr001.Execute SQLQ
        End If
        'Ticket #21677 Franks 09/16/2013 - end
        Exit Sub
    End If
    
    If SaveBGroup <> NewBGroup Then
        If IsWFCUSBenEmp(xEmpNo) Then 'Ticket #23247 Franks 09/16/2013
            Call WFC_UptUSBenByEmp(xEmpNo, CVDate(dlpDate(1).Text), 0, "Y", "Y", , SaveBGroup, dlpDate(1).Text, "Y")
            Exit Sub
        Else
            Msg = "Do you want add/update the Employee's Benefits "
            Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
            If MsgBox(Msg, 36, "info:HR") = 6 Then
                'Call UpdateBenefitGroup
                Call glbUpdateBenefitGroup(xEmpNo, SaveBGroup, NewBGroup, dlpDate(0).Text)
                DoEvents
                glbLEE_ID = xEmpNo
                glbLEE_SName = locSurName
                glbLEE_FName = locFName
                frmBENGRLIST.Show 1
            'Else 'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
            '    If Len(clpBGroup.Text) = 0 Then
            '        SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & xlocEmpnbr
            '        gdbAdoIhr001.Execute SQLQ
            '    End If
            End If
            
            'If the Benefit Group changes, go to the Benefit Group Matrix to update
            'the HREMP 's Coverage Class and Benefit Account on the status/dates screen.
            If Len(NewBGroup) > 0 Then
                SQLQ = "SELECT * FROM HREMP"
                SQLQ = SQLQ & " WHERE HREMP.ED_EMPNBR = '" & xEmpNo & "' "
                rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmp.EOF Then
                    Call getValsFromBenGrpMatrix(NewBGroup, rsEmp("ED_DIV"))
                    If Len(xBenAccount) = 0 Then
                        rsEmp("ED_USER_NUM1") = Null
                    Else
                        rsEmp("ED_USER_NUM1") = xBenAccount
                    End If
                    If Len(xCovClass) = 0 Then
                        rsEmp("ED_USER_TEXT2") = Null
                        rsEmp("ED_USER_TEXT1") = Null 'Ticket #24620 Franks 12/03/2013
                    Else
                        rsEmp("ED_USER_TEXT2") = xCovClass
                        'Ticket #24620 Franks 12/03/2013 - begin
                        If IsNull(rsEmp("ED_USER_TEXT1")) Then rsEmp("ED_USER_TEXT1") = GetBenCertificateNo(NewBGroup, xEmpNo)
                        If Len(rsEmp("ED_USER_TEXT1")) = 0 Then rsEmp("ED_USER_TEXT1") = GetBenCertificateNo(NewBGroup, xEmpNo)
                        'Ticket #24620 Franks 12/03/2013 - end
                    End If
                    rsEmp.Update
                End If
                rsEmp.Close
            End If
        End If
    End If

End Sub

Private Sub getValsFromBenGrpMatrix(NewBGroup, xDIV)
Dim rsBenGrpMrx As New ADODB.Recordset
Dim SQLQ As String
    xCovClass = ""
    xBenAccount = ""
    If Len(NewBGroup) > 0 Then
        SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "' "
        If Len(xDIV) > 0 Then
            SQLQ = SQLQ & "AND BM_DIV = '" & xDIV & "' "
        End If
        rsBenGrpMrx.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsBenGrpMrx.EOF Then
            xCovClass = rsBenGrpMrx("BM_BENEFIT_CLASS")
            xBenAccount = rsBenGrpMrx("BM_BENEFIT_ACCOUNT")
        End If
        rsBenGrpMrx.Close
    End If
    If Len(xBenAccount) = 0 Then txtUserNum1.Text = "" Else txtUserNum1.Text = xBenAccount
    If Len(xCovClass) = 0 Then txtUserText2.Text = "" Else txtUserText2.Text = xCovClass
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
    'Call txtUserText2_Change
    'comUserText2.Text = xCurVal
End Sub

Private Sub txtUserNum1_Change()
If glbWFC Then 'Ticket #22448 Franks
    Call PopcomUserText2(txtUserNum1.Text)
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

Private Sub DispNGSBenGroups() 'Ticket #23247 Franks 09/13/2013
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xLocOrg
Dim xDivCountry
Dim xUnion, xDIV, xStatus

    xDivCountry = GetCountryFromDiv(clpDIV.Text)
    If Not xDivCountry = "U.S.A." Then
        Exit Sub
    End If
    
    xDIV = clpDIV.Text
    If Len(clpCode(10).Text) = 0 Then Exit Sub 'Union is required
    xUnion = clpCode(10).Text
    If Len(clpCode(12).Text) = 0 Then Exit Sub 'Union is required
    xStatus = clpCode(12).Text


    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
    SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
    SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & xStatus & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsTemp.EOF Then 'Ticket #23564 Franks 04/17/2013
    'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" then compare ED_EMP with not equal to
        SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
        SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
        SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
        SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & xStatus & "') " 'convert "-ACT2" to "ACT2"; no "-" then ""
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'if not found then without Status code
        If rsTemp.EOF Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
            SQLQ = SQLQ & "AND ((NG_PLAN_CODE IS NULL) OR NOT( NG_PLAN_CODE ='" & xStatus & "')) "
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        End If
    End If
    
    If Not rsTemp.EOF Then
        clpVadim2.Text = rsTemp("NG_PAY_GROUP")
        'clpVadim1.Text = rsTemp("NG_SUB_GROUP")
        If Not IsNull(rsTemp("NG_BENEFIT_GROUP")) Then 'Ticket #23903 Franks 06/20/2013
            clpBGroup.Text = rsTemp("NG_BENEFIT_GROUP")
        Else
            clpBGroup.Text = ""
        End If
    Else
        'clpVadim1.Text = ""
        clpVadim2.Text = ""
        clpBGroup.Text = "" 'Ticket #23903 Franks 06/20/2013
    End If
    rsTemp.Close
        
End Sub

Private Sub WFCHRSoftDispValues()
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp

'xHRSoftUpt = False
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    SQLQ = SQLQ & "AND SF_UPT_PROCESSED = 0 "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsCanid.EOF Then
    
        'frmTranEMPL.Show 1
        'If glbTran_ID = 0 Then
        '    Unload Me
        '    Exit Sub
        'End If
        'glbTran_ID = rsCanid("SF_EMPNBR")
        glbTran_ID = glbLEE_ID
        
        Screen.MousePointer = HOURGLASS
        
        If Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
            Me.Caption = "Transfer In - " & Left$(glbLEE_SName, 5)
            Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
        End If
        
        lblEENum.Caption = ShowEmpnbr(glbTran_ID)
        
        'If EERetrieve() = False Then Exit Sub
        If EERetrievWFCTran() = False Then Exit Sub
        
        'If Not rsORG.EOF Then
        '    clpDIV = rsORG("TL_NEWDIV")
        'End If
        clpDIV = rsCanid("SF_DIV")
        
        Call WFCScreenSetup 'Ticket #22448

        lblEmpNo = ShowEmpnbr(glbTran_ID)
        
        If Not rsORG.EOF Then
            dlpDate(0) = rsCanid("SF_STARTDATE")
            dlpDate(1) = dlpDate(0)
            'Ticket #21677 Franks 03/14/2012
            If Not IsNull(rsCanid("SF_ORG")) Then
                clpCode(10).Text = rsCanid("SF_ORG")
                Call WFC_UnionScreen(rsCanid("SF_ORG"))
                Call getPayGroup(rsCanid("SF_DIV"), rsCanid("SF_ORG"))
            End If
            
            'employee data from hrsoft
            If Not IsNull(rsCanid("SF_SECTION")) Then clpCode(0).Text = rsCanid("SF_SECTION")
            If Not IsNull(rsCanid("SF_LOC")) Then clpCode(7).Text = rsCanid("SF_LOC")
            If Not IsNull(rsCanid("SF_REGION")) Then clpCode(6).Text = rsCanid("SF_REGION")
            If Not IsNull(rsCanid("SF_ADMINBY")) Then clpCode(5).Text = rsCanid("SF_ADMINBY")
            If Not IsNull(rsCanid("SF_EMPCODE")) Then clpCode(12).Text = rsCanid("SF_EMPCODE")
            Call clpCode_LostFocus(12)
            
            clpGLNum.TextBoxWidth = 1200
            'Call Set_PositionSalary
            Call Set_PositionSalaWFCHRSoft(rsCanid)
            If glbWFC Then
                ODateOfHire = dlpDate(3).Text
            End If
            xUpdateable = True

            MDIMain.panHelp(1).Caption = " "
    
            Call INI_Controls(Me)
            'xHRSoftUpt = True
        End If
        
        'glbTran_ID
        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbTran_ID & " "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp("ER_OTHERDATE1")) Then
                dlpDate(6).Text = rsTemp("ER_OTHERDATE1")
            End If
        End If
        rsTemp.Close
        
        Call WFC_DispNGSStartDate 'Ticket #24652 Franks 12/02/2013
    End If
End If
End Sub

Private Sub WFC_DispNGSStartDate() 'Ticket #24652 Franks 12/02/2013
If glbWFC Then
    If lblNGSStart.Visible Then
        If glbCandidate > 0 And Not (OldUnion = clpCode(10).Text) Then
        'Ticket #24767 Franks 12/11/2013
        'Talked this with Jerry on 12/09/2013, he said the logic below works for info:HR, but HRsoft can change both Div and Union at the same time,
        'so we will add more logic for HRsoft import only: if Union or both div and union changed then ngs start date = the Transfer In Date
                If IsDate(dlpDate(1).Text) Then
                    dlpDate(6).Text = dlpDate(1).Text
                End If
        Else
            If OldDiv = clpDIV.Text And Not (OldUnion = clpCode(10).Text) Then
                If IsDate(dlpDate(1).Text) Then
                    dlpDate(6).Text = dlpDate(1).Text
                End If
            Else
                If Len(dlpDate(6).Text) = 0 Then 'Ticket #24620 Franks 12/03/2013 'no old NGS then default to Transfer In data
                    If IsDate(dlpDate(1).Text) Then
                        dlpDate(6).Text = dlpDate(1).Text
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub WFCDivTranInSamePlant() 'Ticket #25221 Franks 03/17/2014
Dim rs As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM LN_TRALOG WHERE TL_TERM_SEQ = " & glbTran_Seq & " "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        glbTran_ID = rs("TL_EMPNBR")
        'glbTran_Seq = rs("TL_TERM_SEQ")
        glbTERM_Seq = glbTran_Seq
        glbTERM_ID = glbTran_ID
        If Not IsNull(rs("TL_FNAME")) Then
            glbTran_Fname = rs("TL_FNAME")
        Else
            glbTran_Fname = "*ERROR*"
        End If
        If Not IsNull(rs("TL_SURNAME")) Then
            glbTran_SName = rs("TL_SURNAME")
        Else
            glbTran_SName = "*ERROR*"
        End If
    End If
    rs.Close
End Sub

Private Sub WFC_DivTranEmpPosSalary() 'Ticket #25221 Franks 03/17/2014
Dim EMP_Snap As New ADODB.Recordset
Dim HRSH_Snap As New ADODB.Recordset
Dim HRJH_Snap As New ADODB.Recordset
Dim SQLQ, xDATE, xLinDate, NewDate, dtY1%, dtY2%, xSalary, xSalCD


SQLQ = "SELECT * FROM TERM_HREMP WHERE TERM_SEQ=" & fglbTERM_Seq
EMP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not (EMP_Snap.BOF And EMP_Snap.EOF) Then
    xLinDate = EMP_Snap("ED_DOH")
    If Len(clpCode(1).Text) = 0 Then clpCode(1).Text = "DEPT" 'default???
    '---- employee info begin -------------
    If Not IsNull(EMP_Snap("ED_PAYROLL_ID")) Then txtPayrollID.Text = EMP_Snap("ED_PAYROLL_ID")
    If Not IsNull(EMP_Snap("ED_DEPTNO")) Then clpDept.Text = EMP_Snap("ED_DEPTNO")
    If Not IsNull(EMP_Snap("ED_GLNO")) Then clpGLNum.Text = EMP_Snap("ED_GLNO")
    If Not IsNull(EMP_Snap("ED_LOC")) Then clpCode(7).Text = EMP_Snap("ED_LOC")
    If Not IsNull(EMP_Snap("ED_SECTION")) Then clpCode(0).Text = EMP_Snap("ED_SECTION")
    If Not IsNull(EMP_Snap("ED_REGION")) Then clpCode(6).Text = EMP_Snap("ED_REGION")
    If Not IsNull(EMP_Snap("ED_DOH")) Then dlpDate(3) = EMP_Snap("ED_DOH")
    If Not IsNull(EMP_Snap("ED_SENDTE")) Then dlpDate(4) = EMP_Snap("ED_SENDTE")
    If Not IsNull(EMP_Snap("ED_LTHIRE")) Then dlpDate(5) = EMP_Snap("ED_LTHIRE")
    If Not IsNull(EMP_Snap("ED_EMPTYPE")) Then
        If EMP_Snap("ED_EMPTYPE") = "Y" Then
            comEmpType.ListIndex = 0
        End If
        If EMP_Snap("ED_EMPTYPE") = "N" Then
            comEmpType.ListIndex = 1
        End If
        txtEmpType.Text = EMP_Snap("ED_EMPTYPE")
    End If
    If Not IsNull(EMP_Snap("ED_EMP")) Then clpCode(12).Text = EMP_Snap("ED_EMP")
    'If Not IsNull(EMP_Snap("ED_ORG")) Then clpCode(10).Text = EMP_Snap("ED_ORG")
    If Not IsNull(EMP_Snap("ED_BENEFIT_GROUP")) Then clpBGroup.Text = EMP_Snap("ED_BENEFIT_GROUP")
    If Not IsNull(EMP_Snap("ED_USER_NUM1")) Then txtUserNum1.Text = EMP_Snap("ED_USER_NUM1") 'Benefit Account
    If Not IsNull(EMP_Snap("ED_USER_TEXT2")) Then 'Coverage Class
        txtUserText2.Text = EMP_Snap("ED_USER_TEXT2")
        comUserText2.ListIndex = FindCBIndex(comUserText2, Left((txtUserText2 & " - "), 4), 4)
    End If
    If Not IsNull(EMP_Snap("ED_VADIM2")) Then clpVadim2.Text = EMP_Snap("ED_VADIM2")
    '---- employee info end -------------

    'If IsDate(xLinDate) Then
    '-------------Salary begin -------------
    SQLQ = "Select * from TERM_SALARY_HISTORY WHERE TERM_SEQ=" & fglbTERM_Seq & " AND SH_CURRENT <>0"
    HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xDATE = 0: xSalary = "": xSalCD = ""
    If Not (HRSH_Snap.BOF And HRSH_Snap.EOF) Then
        xDATE = HRSH_Snap("SH_EDATE")
        xSalary = HRSH_Snap("SH_SALARY")
        xSalCD = HRSH_Snap("SH_SALCD")
        medSalary.Text = xSalary
        comPayPer.ListIndex = IIf(xSalCD = "", -1, IIf(xSalCD = "A", 0, IIf(xSalCD = "H", 1, IIf(xSalCD = "M", 2, 3))))
        lblSalCode = xSalCD
        
        If Not IsNull(HRSH_Snap("SH_NEXTDAT")) Then dlpDate(2).Text = HRSH_Snap("SH_NEXTDAT")
        If glbWFC Then
            If Not IsNull(HRSH_Snap("SH_FISCALYEAR")) Then
                txtFiscalYear.Text = HRSH_Snap("SH_FISCALYEAR")
            End If
            If Not IsNull(HRSH_Snap("SH_MARKETLINE")) Then
                txtMarketLine.Text = HRSH_Snap("SH_MARKETLINE")
            End If
            If Not IsNull(HRSH_Snap("SH_PAYP")) Then 'Ticket #15818
                clpCode(9).Text = HRSH_Snap("SH_PAYP")
            End If
            Call Set_MarketLine_List
        End If

    End If
    HRSH_Snap.Close
    '-------------Salary end -------------
        
    '-------------Position begin -------------
    SQLQ = "Select * from TERM_JOB_HISTORY WHERE TERM_SEQ=" & fglbTERM_Seq & " AND JH_CURRENT <>0"
    HRJH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not HRJH_Snap.EOF Then

        clpJob = HRJH_Snap("JH_JOB")
        If Not IsNull(HRJH_Snap("JH_DHRS")) Then medHours(0).Text = HRJH_Snap("JH_DHRS")
        If Not IsNull(HRJH_Snap("JH_WHRS")) Then medHours(1).Text = HRJH_Snap("JH_WHRS")
        If Not IsNull(HRJH_Snap("JH_PHRS")) Then medHours(2).Text = HRJH_Snap("JH_PHRS")
        If Not IsNull(HRJH_Snap("JH_REPTAU")) Then elpReptAuthShow(0).Text = HRJH_Snap("JH_REPTAU")

    End If
    HRJH_Snap.Close
    '-------------Position end -------------
End If
EMP_Snap.Close

End Sub

Private Sub DelTermEEO(EESEQ&)
Dim SQLQ As String
    SQLQ = "DELETE FROM Term_HREEO WHERE (Term_HREEO.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001X.Execute SQLQ
End Sub

Private Function getTermEmpPayID()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
Dim retVal
    retVal = ""
    SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID, TERM_SEQ FROM TERM_HREMP WHERE TERM_SEQ=" & fglbTERM_Seq
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
            retVal = rsEmp("ED_PAYROLL_ID")
        End If
    End If
    rsEmp.Close
    getTermEmpPayID = retVal
End Function
