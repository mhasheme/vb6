VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRTSStatus 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Timesheet Status Report"
   ClientHeight    =   9030
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   10380
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
   ScaleHeight     =   9030
   ScaleWidth      =   10380
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWeek 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2265
      TabIndex        =   9
      Top             =   3450
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2265
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   19
      Tag             =   "00-Shift"
      Top             =   5460
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1950
      TabIndex        =   18
      Tag             =   "EDSE-Section "
      Top             =   5130
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1950
      TabIndex        =   17
      Tag             =   "EDAB-Administered By"
      Top             =   4800
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1950
      TabIndex        =   16
      Tag             =   "EDRG-Region"
      Top             =   4470
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1950
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2115
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1950
      TabIndex        =   4
      Top             =   1785
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   2
      Tag             =   "EDLC-Location"
      Top             =   1140
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   810
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   480
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1950
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2445
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   8280
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1950
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1455
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3570
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "40-Date from and including this date forward"
      Top             =   3780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin Threed.SSCheck chkNotEntered 
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Tag             =   "Not Entered"
      Top             =   6240
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Not Entered"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkSaved 
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Tag             =   "SAVED"
      Top             =   6600
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Saved"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkApproved 
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Tag             =   "APPROVED"
      Top             =   6960
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Approved"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkRejected 
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Tag             =   "REJECTED"
      Top             =   6600
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Rejected"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkSubmitted 
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Tag             =   "SUBMITTED"
      Top             =   7320
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Submitted"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   3
      Left            =   8460
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   2
      Left            =   6840
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "40-Date from and including this date forward"
      Top             =   3780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtWeekTo 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7155
      TabIndex        =   10
      Top             =   3450
      Width           =   1335
   End
   Begin Threed.SSCheck chkReSubmitted 
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Tag             =   "RESUBMITTED"
      Top             =   6240
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Resubmitted"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpSUP 
      Height          =   285
      Index           =   1
      Left            =   1950
      TabIndex        =   7
      Tag             =   "00-Employee Number "
      Top             =   2775
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPAYP 
      DataField       =   "PP_PAYP"
      Height          =   285
      Left            =   1950
      TabIndex        =   15
      Tag             =   "SDPP-Pay Period"
      Top             =   4110
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin Threed.SSCheck chkAppFwd 
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Tag             =   "APP/FWD"
      Top             =   6960
      Width           =   3315
      _Version        =   65536
      _ExtentX        =   5847
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Timesheet Approve/Forward"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period Code"
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
      Left            =   150
      TabIndex        =   46
      Top             =   4155
      Width           =   1185
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "AttSupervisor"
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
      Left            =   150
      TabIndex        =   45
      Top             =   2820
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Timesheet Status :"
      Height          =   195
      Left            =   180
      TabIndex        =   44
      Top             =   6270
      Width           =   1755
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Left            =   5280
      TabIndex        =   43
      Top             =   3825
      Width           =   1095
   End
   Begin VB.Label lblWeekTo 
      Caption         =   "To Pay Period #"
      Height          =   195
      Left            =   5280
      TabIndex        =   42
      Top             =   3495
      Width           =   1395
   End
   Begin VB.Image imgIconTo 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6810
      Picture         =   "FZTSStatusRpt.frx":0000
      Top             =   3465
      Width           =   240
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Left            =   150
      TabIndex        =   41
      Top             =   3825
      Width           =   1095
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1680
      Picture         =   "FZTSStatusRpt.frx":014A
      Top             =   3465
      Width           =   240
   End
   Begin VB.Label lblWeek 
      Caption         =   "From Pay Period #"
      Height          =   195
      Left            =   150
      TabIndex        =   40
      Top             =   3495
      Width           =   1395
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Left            =   150
      TabIndex        =   39
      Top             =   3165
      Width           =   1395
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   150
      TabIndex        =   38
      Top             =   5505
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      Left            =   150
      TabIndex        =   37
      Top             =   1500
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   150
      TabIndex        =   36
      Top             =   1830
      Width           =   450
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
      Left            =   150
      TabIndex        =   35
      Top             =   1185
      Width           =   615
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      Left            =   150
      TabIndex        =   34
      Top             =   4845
      Width           =   1125
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   150
      TabIndex        =   33
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      Left            =   150
      TabIndex        =   32
      Top             =   5175
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
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
      Left            =   150
      TabIndex        =   31
      Top             =   4515
      Width           =   510
   End
   Begin VB.Label lblSelCri 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   150
      Width           =   1575
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Index           =   0
      Left            =   150
      TabIndex        =   29
      Top             =   2490
      Width           =   1290
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   150
      TabIndex        =   28
      Top             =   855
      Width           =   825
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   27
      Top             =   525
      Width           =   555
   End
End
Attribute VB_Name = "frmRTSStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsnapEENames As Recordset
Dim DATE1, DATE2

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%
On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm(Me.Caption, Me) Then Exit Sub
    Call set_PrintState(False)
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
     Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet Status", "Timesheet Status Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_REGION"
    Case 2: strCd$ = "HREMP.ED_SECTION"
    Case 3: strCd$ = "HREMP.ED_EMP"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_ORG"
    Case 6: strCd$ = "HREMP.ED_PT"
    End Select
    'CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "((" & strCd$ & " = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or (" & strCd$ & " = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "HREMP.ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$
    Cri_SetAll = False
    '''On Error GoTo modSetCriteria_Err
    Screen.MousePointer = HOURGLASS
    
    glbiOneWhere = False
    glbstrSelCri = ""
    
    Call glbCri_DeptUN(clpDept.Text)
    'Call Cri_Dept
    Call Cri_Div
    For X% = 0 To 6
        Call Cri_Code(X%)
    Next X%
    Call Cri_EE
    'Call AttWrk
    Call Export_Timesheet_Status

    'strRName$ = glbIHRREPORTS & "rztimesheet.rpt"
    
    'Me.vbxCrystal.ReportFileName = strRName$
    ''x% = Cri_Sorts()   ' returns number of sections formated
    'Me.vbxCrystal.SelectionFormula = "{HR_ATT_TIMESHEET.AD_WRKEMP}='" & glbUserID & "'"
    'Me.vbxCrystal.WindowTitle = Me.Caption
    'If glbSQL Or glbOracle Then
    '    Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = RptODBC_SQL
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    '    'For x% = 1 To 7
    '    '    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    '    'Next x%
    'End If
    
    Cri_SetAll = True

    Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet Status", "Timesheet Status Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function CriCheck()
Dim X%

CriCheck = False

If Trim(txtYear) = "" Then
    MsgBox "Pay Period Year cannot be blank"
    txtYear.SetFocus
    Exit Function
Else
    If Val(txtYear) > Year(Date) + 100 Or Val(txtYear) < Year(Date) - 100 Then
        MsgBox "Pay Period Year is not a valid year"
        txtYear.SetFocus
        Exit Function
    End If
End If

If Len(txtWeek.Text) = 0 Then
    MsgBox "From Pay Period # is requried field"
    txtWeek.SetFocus
    Exit Function
End If

If Not IsNumeric(txtWeek.Text) Or InStr(txtWeek.Text, ".") > 0 Or txtWeek.Text = "0" Then
    MsgBox "Invalid From Pay Period #"
    txtWeek.SetFocus
    Exit Function
End If

If Len(txtWeekTo.Text) = 0 Then
    MsgBox "To Pay Period # is requried field"
    txtWeekTo.SetFocus
    Exit Function
End If

If Not IsNumeric(txtWeekTo.Text) Or InStr(txtWeekTo.Text, ".") > 0 Or txtWeekTo.Text = "0" Then
    MsgBox "Invalid To Pay Period #"
    txtWeekTo.SetFocus
    Exit Function
End If

If IsNumeric(txtWeek.Text) And IsNumeric(txtWeekTo.Text) Then
    If CInt(txtWeek.Text) > CInt(txtWeekTo.Text) Then
        MsgBox "From Pay Period # cannot be greater than To Pay Period #"
        txtWeek.SetFocus
        Exit Function
    End If
End If

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not chkNotEntered And Not chkSaved And Not chkApproved And Not chkSubmitted And Not chkRejected And Not chkReSubmitted Then
    MsgBox "At least one Timesheet Status has to be selected"
    chkNotEntered.SetFocus
    Exit Function
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Screen.MousePointer = HOURGLASS
glbOnTop = Me.name

'Ticket #28955 - Commented these label as it is messing up the labels being called again below in 'setRptCaption(Me)'
'Call setCaption(lblDiv)
'Call setCaption(lblRegion)
'Call setCaption(lblSection)
'Call setCaption(lblDept)
Call setCaption(lblEENum(1))
Call setRptCaption(Me)
'lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

If lblEENum(1).Caption = "AttSupervisor" Then lblEENum(1).Caption = "Supervisor"

If glbCompSerial = "S/N - 2227W" Then clpCode(1).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

txtYear = Year(Date)

'chkShowEmp.Visible = True
If glbLinamar Then
    clpCode(1).MaxLength = 8
End If
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
gdbAdoIhr001.Execute "DELETE FROM HR_ATT_TIMESHEET " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub Cri_Dept()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String

'Ticket #21968 - Allow multi code selection criteria
'If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO = '" & clpDept.Text & "') "
If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO in ['" & Replace(clpDept.Text, ",", "','") & "']) "

glbstrSelCri = glbSeleDeptUn & DeptCri
End Sub

Private Sub imgIconTo_Click()
    frmPayPeriodList.SelectedYear = Val(txtYear)
    'frmPayPeriodList.PayPeriodCode = clpPayP.Text
    frmPayPeriodList.Show 1
    txtWeekTo = glbWeek
    dlpDateRange(2) = glbFrom
    dlpDateRange(3) = glbTo
    'Ticket #29984 - Receiving Pay Period code as well if Pay Period list is by Pay Period Code
    clpPayP.Text = glbPayP
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub txtWeekTo_Change()
    Dim DateRange
    
    DateRange = Split(getDateRange("", txtWeekTo, txtYear), "|")
    dlpDateRange(2) = DateRange(0)
    dlpDateRange(3) = DateRange(1)
End Sub

Private Sub txtWeekTo_DblClick()
    Call imgIconTo_Click
End Sub

Private Sub txtWeekTo_LostFocus()
    If txtWeekTo = "" Then
        dlpDateRange(2) = ""
        dlpDateRange(3) = ""
    Else
        'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
    End If
End Sub

Private Sub txtYear_GotFocus()      ' Serbo
Call SetPanHelp(Me.ActiveControl)   '
End Sub                             '

Function StripChar(StringToStrip, CharToStrip)
    Dim I, buf, OneChar
    
    For I = 1 To Len(StringToStrip)
        OneChar = Mid(StringToStrip, I, 1)
        If OneChar <> CharToStrip Then buf = buf & OneChar
    Next I
    StripChar = buf
End Function

Private Sub txtWeek_Change()
    Dim DateRange
    
    DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Private Sub txtWeek_DblClick()
    Call imgIcon_Click
End Sub

Private Sub imgIcon_Click()
    frmPayPeriodList.SelectedYear = Val(txtYear)
    'frmPayPeriodList.PayPeriodCode = clpPayP.Text
    frmPayPeriodList.Show 1
    txtWeek = glbWeek
    dlpDateRange(0) = glbFrom
    dlpDateRange(1) = glbTo
    'Ticket #29984 - Receiving Pay Period code as well if Pay Period list is by Pay Period Code
    clpPayP.Text = glbPayP
End Sub

Private Sub txtWeek_LostFocus()
    If txtWeek = "" Then
        dlpDateRange(0) = ""
        dlpDateRange(1) = ""
    Else
        'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
    End If
End Sub

Private Sub txtYear_Change()
    Dim DateRange
    
    DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
    
    DateRange = Split(getDateRange("", txtWeekTo, txtYear), "|")
    dlpDateRange(2) = DateRange(0)
    dlpDateRange(3) = DateRange(1)
    
End Sub

Function getDateRange(theClientNumber, thePayNbr, theYear)
    Dim rsPayPeriod As New ADODB.Recordset
    Dim SQLQ, intNum
    
    On Error Resume Next
    
    getDateRange = "|"
    
    If Not IsNumeric(thePayNbr) Then Exit Function
    If Not IsNumeric(theYear) Then Exit Function
    
    SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_PAYPERIOD WHERE PP_PAYP='" & theClientNumber & "'"
    SQLQ = SQLQ & " and PP_NBR = " & thePayNbr
    SQLQ = SQLQ & " and PP_YEAR = '" & theYear & "'"
    rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsPayPeriod.EOF Then
        getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
    End If
    rsPayPeriod.Close
    Exit Function

End Function

Private Function getWSQLQ(WithAtt As Boolean)
Dim QStr
QStr = glbSeleDeptUn
If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
If clpCode(1) <> "" Then QStr = QStr & " AND ED_ORG in ('" & Replace(clpCode(1), ",", "','") & "')"
If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpCode(6) <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpCode(6), ",", "','") & "')"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If WithAtt Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND AD_DOA>=" & Date_SQL(DATE1)
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND AD_DOA<=" & Date_SQL(DATE1)
'    If clpAtt <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpAtt, ",", "','") & "') "
'    If clpPayP <> "" Then QStr = QStr & " AND ES_CRSCODE IN ('" & Replace(clpPayP, ",", "','") & "') "
'    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
'        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
'        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
'        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
'        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
'        QStr = QStr & ")"
'    End If
End If

getWSQLQ = QStr
End Function

Private Sub AttWrk()
'Dim CoJobCodeS As New Collection
'Dim CoJobCode As New Collection
'Dim rsHours As New ADODB.Recordset
Dim rsAT As New ADODB.Recordset
Dim rsATTCal As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strStatus

Dim iWorksheet As Integer
Dim iDate, DayNum As Integer
Dim strRANGE As String
Dim sFlag As Boolean
Dim xRegularRate
Dim dtDate As Date
Dim strTableName
'Dim blnFound
Dim xDay, xField
Dim xHours
Dim SQLQ
Dim xEmpnbr
Dim gdbESS As New ADODB.Connection

If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If

On Error GoTo Err_XLS
    DATE1 = dlpDateRange(0)
    DATE2 = dlpDateRange(1)
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 20
    MDIMain.panHelp(1).Caption = " Please Wait"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_ATT_TIMESHEET " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.CommitTrans
    
    If Not glbSQL And Not glbOracle Then Pause (0.5)
    
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE "
    SQLQ = SQLQ & getWSQLQ(False)

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsTemp.EOF Then
         MsgBox ("There is no employee based on the search criteria.")
    End If
     
    Do While Not rsTemp.EOF  'processing all employees who have timesheet data in HR_Timesheet or HR_Attendance
        xEmpnbr = rsTemp.Fields("ED_EMPNBR")
        strStatus = getStatus(xEmpnbr, DATE1, DATE2)
        Select Case strStatus
        Case "APPROVED"
            strTableName = "HR_ATTENDANCE"
        Case ""
            'strTableName = "HR_ATTENDANCE"
            GoTo Loopend
        Case Else
            strTableName = "HR_TIMESHEET"
        End Select

        SQLQ = "SELECT AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,AD_EMPNBR,AD_DOA,AD_HRS,AD_REASON,  AD_SHIFT "
        SQLQ = SQLQ & " FROM " & strTableName
        '& ", HRTABL "
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(DATE1)
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(DATE2)
        'SQLQ = SQLQ & " AND (AD_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
        
'        If strTableName = "HR_ATTENDANCE" Then
'            SQLQ = SQLQ & " UNION "
'            SQLQ = SQLQ & " SELECT AH_COMPNO,'" & glbUserID & "' AS AH_WRKEMP,AH_EMPNBR,AH_DOA,AH_HRS,AH_REASON,  AH_SHIFT "
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HRTABL "
'            SQLQ = SQLQ & " WHERE AH_EMPNBR =" & rsTemp.Fields("ED_EMPNBR")
'            SQLQ = SQLQ & " AND AH_DOA>=" & Date_SQL(DATE1)
'            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(DATE2)
'            SQLQ = SQLQ & " AND (AH_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
'
'        End If
        SQLQ = SQLQ & " order by AD_DOA"
        
        If glbSQL Or glbOracle Then
            rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
        Else
            If strTableName = "HR_TIMESHEET" Then
                rsAT.Open SQLQ, gdbESS, adOpenKeyset, adLockReadOnly
            Else
                rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
            End If
        End If
        '        xxx = rsAT.RecordCount
        '        xx1 = 0
        Do Until rsAT.EOF
'            xx1 = xx1 + 1
'            MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 60 + 30
            xDay = DateDiff("d", DATE1, rsAT!AD_DOA) + 1
            xField = "AD_DAY" & xDay
            
            SQLQ = "Select * from HR_ATT_TIMESHEET  "
            SQLQ = SQLQ & " where AD_EMPNBR=" & xEmpnbr
            SQLQ = SQLQ & " AND  AD_WRKEMP='" & glbUserID & "'"
            SQLQ = SQLQ & " AND AD_REASON ='" & rsAT!AD_REASON & "'"
            
            If glbSQL Or glbOracle Then
                rsATTCal.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            Else
                rsATTCal.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
            End If
            
            xHours = 0
            If rsATTCal.EOF Then
                rsATTCal.AddNew
                rsATTCal!AD_EMPNBR = xEmpnbr
                rsATTCal!AD_COMPNO = "001"
                rsATTCal!AD_WRKEMP = glbUserID
                rsATTCal!AD_DOA = DATE1
                rsATTCal!AD_REASON = rsAT!AD_REASON
            End If
            rsATTCal!AD_STATUS = strStatus
            If IsNull(rsATTCal(xField)) Then
                xHours = 0
            Else
                xHours = rsATTCal(xField)
            End If
            rsATTCal(xField) = xHours + rsAT!AD_HRS
            rsATTCal.Update
            rsATTCal.Close
            rsAT.MoveNext
        Loop
        rsAT.Close
Loopend:
        rsTemp.MoveNext
    Loop
    rsTemp.Close

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Exit Sub
Err_XLS:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT

    If Err = 1004 Then
        Resume Next
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet", "", "Select")
Resume Next
End Sub

Private Function getStatus(xEmpnbr, strStartDate, strEndDate)
Dim SQLQ, statusFlag
Dim rsDS As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection

    If glbSQL Or glbOracle Then
        Set gdbESS = gdbAdoIhr001
    Else
        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
    End If
    
    On Error Resume Next
    SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_TIMESHEET "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(strStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(strEndDate)

    If glbSQL Or glbOracle Then
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsDS.Open SQLQ, gdbESS, adOpenForwardOnly
    End If
    
    getStatus = ""
    statusFlag = True
    Do While Not rsDS.EOF
        If statusFlag Then
            If IsNull(rsDS("AD_APPROVED")) Then
                If rsDS("AD_UPLOAD") & "" = "Y" Then
                    getStatus = "SUBMITTED"
                Else
                    getStatus = "SAVED"
                End If
            Else
                getStatus = rsDS("AD_APPROVED")
                If getStatus = "RESUBMIT" Then getStatus = "RESUBMITTED"
            End If
            statusFlag = False
        Else
            'getStatus="Inconsistent"
            getStatus = "SAVED"
        End If
        rsDS.MoveNext
    Loop
    rsDS.Close
    
    If getStatus = "" Then
        SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(strStartDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(strEndDate)
    
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        statusFlag = True
        If Not rsDS.EOF Then
            getStatus = "APPROVED"
        End If
        rsDS.Close
    
    End If
    Set rsDS = Nothing
    If Err.Number <> 0 Then
    End If
End Function

Function getCompTimeBank(xEmpnbr)
    Dim SQLQ, xOTEarned, xCTTaken
    Dim rsAttOT As New ADODB.Recordset
    Dim rsAttCT As New ADODB.Recordset
    SQLQ = ""
    On Error Resume Next
    getCompTimeBank = 0
    
    'EARNED - OT
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS AD_HRSTOTAL FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    
    If glbOracle Then
        SQLQ = SQLQ & " AND SUBSTR(AD_REASON,1,2) = 'OT'"
    Else
        SQLQ = SQLQ & " AND LEFT(AD_REASON,2) = 'OT'"
    End If
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    rsAttOT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    xOTEarned = 0
    If rsAttOT.EOF Then
        xOTEarned = 0
    Else
        If Not IsNull(rsAttOT("AD_HRSTOTAL")) Then xOTEarned = CDbl(rsAttOT("AD_HRSTOTAL"))
    End If
    rsAttOT.Close
    
    'TAKEN - CT
    SQLQ = ""
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS AD_HRSTOTAL FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    
    If glbOracle Then
        SQLQ = SQLQ & " AND SUBSTR(AD_REASON,1,2) = 'CT'"
    Else
        SQLQ = SQLQ & " AND LEFT(AD_REASON,2) = 'CT'"
    End If
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    rsAttCT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    xCTTaken = 0
    If rsAttCT.EOF Then
        xCTTaken = 0
    Else
        If Not IsNull(rsAttCT("AD_HRSTOTAL")) Then xCTTaken = CDbl(rsAttCT("AD_HRSTOTAL"))
    End If
    rsAttCT.Close
    
    getCompTimeBank = FormatNumber((xOTEarned - xCTTaken), 2)
    
    If Err.Number <> 0 Then
    End If
End Function

Function getVacOutstDay(xEmpnbr)
    Dim SQLQ, xOuts
    Dim rsEmp As New ADODB.Recordset
    SQLQ = ""
    On Error Resume Next
    
    getVacOutstDay = 0
    
    SQLQ = "SELECT ED_EFDATE,ED_ETDATE, ED_VAC,ED_PVAC,ED_VACT,ED_DHRS FROM HREMP "
    SQLQ = SQLQ & " WHERE ED_EMPNBR =" & xEmpnbr
    
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsEmp.EOF Then Exit Function

    xOuts = 0
    If Not IsNull(rsEmp("ED_VAC")) Then xOuts = xOuts + CDbl(rsEmp("ED_VAC"))
    If Not IsNull(rsEmp("ED_PVAC")) Then xOuts = xOuts + CDbl(rsEmp("ED_PVAC"))
    If Not IsNull(rsEmp("ED_VACT")) Then xOuts = xOuts - CDbl(rsEmp("ED_VACT"))
    If Not IsNull(rsEmp("ED_DHRS")) Then
        xOuts = FormatNumber(xOuts / CDbl(rsEmp("ED_DHRS")), 2)
    Else
        xOuts = 0
    End If
    getVacOutstDay = xOuts
    
    rsEmp.Close
    If Err.Number <> 0 Then
    End If
End Function

Function getSickTaken(xEmpnbr)
    Dim SQLQ, xTaken
    Dim rsEmp As New ADODB.Recordset
    
    SQLQ = ""
    On Error Resume Next
    getSickTaken = 0
    SQLQ = "SELECT ED_EFDATES,ED_ETDATES,ED_SICK,ED_PSICK,ED_SICKT,ED_DHRS FROM HREMP "
    SQLQ = SQLQ & " WHERE ED_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND ED_EFDATES <=" & Date_SQL(dlpDateRange(1))
    SQLQ = SQLQ & " AND ED_ETDATES >=" & Date_SQL(dlpDateRange(0))
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsEmp.EOF Then Exit Function

    xTaken = 0
    If Not IsNull(rsEmp("ED_PSICK")) Then xTaken = xTaken + CDbl(rsEmp("ED_PSICK"))
    If Not IsNull(rsEmp("ED_SICK")) Then xTaken = xTaken + CDbl(rsEmp("ED_SICK"))
    If Not IsNull(rsEmp("ED_SICKT")) Then xTaken = xTaken - CDbl(rsEmp("ED_SICKT"))
    If Not IsNull(rsEmp("ED_DHRS")) Then
        xTaken = FormatNumber(xTaken / CDbl(rsEmp("ED_DHRS")), 2)
    Else
        xTaken = 0
    End If
    getSickTaken = xTaken
        
    rsEmp.Close
    
    If Err.Number <> 0 Then
    End If
End Function

Function getHrsTaken(xEmpnbr)
    Dim SQLQ2, SQLQ1
    Dim rsHour As New ADODB.Recordset
    Dim rsTbl As New ADODB.Recordset
    Dim xNo, xHDesc, xTaken
    xNo = 0
    On Error Resume Next
    
    
    SQLQ2 = ""
    SQLQ2 = "SELECT * from HRENTHRS "
    SQLQ2 = SQLQ2 & " WHERE HE_EMPNBR = " & xEmpnbr
    SQLQ2 = SQLQ2 & " AND HE_FDATE <=" & Date_SQL(dlpDateRange(1))
    SQLQ2 = SQLQ2 & " AND HE_TDATE >=" & Date_SQL(dlpDateRange(0))
    SQLQ2 = SQLQ2 & " ORDER BY HE_FDATE DESC"
    
    rsHour.Open SQLQ2, gdbAdoIhr001, , adOpenForwardOnly

    If rsHour.EOF Then Exit Function

    rsHour.MoveFirst
    
    getHrsTaken = ""
    Do While Not rsHour.EOF
        If Not (glbCompSerial = "S/N - 2257W" And (Left(rsHour("HE_TYPE"), 2) = "BT" Or Left(rsHour("HE_TYPE"), 3) = "MA2")) Then 'not Hamilton CCAS And Reason Code not is BT or MA2
            SQLQ1 = ""
            SQLQ1 = "SELECT TB_KEY,TB_DESC FROM HRTABL "
            SQLQ1 = SQLQ1 & " WHERE (TB_NAME='ADRE') "
            SQLQ1 = SQLQ1 & " AND (TB_KEY='" & rsHour("HE_TYPE") & "')"
            rsTbl.Open SQLQ1, gdbAdoIhr001, , adOpenForwardOnly

            If rsTbl.EOF Then Exit Function

            xHDesc = ""
            If Not IsNull(rsTbl("TB_DESC")) Then xHDesc = xHDesc & rsTbl("TB_DESC") & " Remaining"
            getHrsTaken = getHrsTaken & xHDesc & "|"
    
            xTaken = 0
            If Not IsNull(rsHour("HE_ENTITLE")) Then xTaken = xTaken + CDbl(rsHour("HE_ENTITLE"))
            If Not IsNull(rsHour("HE_TAKEN")) Then xTaken = xTaken - CDbl(rsHour("HE_TAKEN"))
            getHrsTaken = getHrsTaken & FormatNumber(xTaken, 2) & "|"

            rsTbl.Close
        
            xNo = xNo + 1
        End If
        rsHour.MoveNext
    Loop
        
    rsHour.Close

    getHrsTaken = xNo & "|" & getHrsTaken & "|"
    
End Function

Private Sub Export_Timesheet_Status()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsPayPrd As New ADODB.Recordset
    Dim rsCompInfo As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xStatus As String
    Dim xRow, xCol As Long
    Dim I, totNum, xEmpDisp
    Dim noStatus As Boolean
    Dim xExcelRptPath  As String

    On Error GoTo Export_Timesheet_Status_Err
    
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If

    'Get Employees to display
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP "
    SQLQ = SQLQ & " WHERE " & sSQLQ
    
    If Len(Trim(elpSUP(1).Text)) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ")) "
    End If
    
    'Ticket #29984 - Filter employee by Pay Period Code
    If Len(clpPayP.Text) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT SH_EMPNBR FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_PAYP in ('" & Replace(clpPayP.Text, ",", "','") & "'))"
    End If
    
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME, ED_EMPNBR"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        xEmpDisp = 0
        
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TSStatusTmp.xls"
        
        'Ticket #22034 - May need to save the report in different path
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TSStatusRpt" & Trim(glbUserID) & ".xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "TSStatusRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        SQLQ = "select PC_NAME from hrparco"
        rsCompInfo.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Dim companyname As String
        If Not rsCompInfo.EOF Then
            rsCompInfo.MoveFirst
            companyname = rsCompInfo("PC_NAME")
        End If
    
        rsCompInfo.Close
        
        exSheet.Cells(1, 3) = companyname
        
        exSheet.Cells(1, 2) = Format(Now, "mm/dd/yyyy")
        exSheet.Cells(2, 2) = Time$
        
        'Display Pay Periods - Column Headings
        SQLQ = "SELECT * FROM HR_PAYPERIOD "
        SQLQ = SQLQ & " WHERE PP_YEAR =" & txtYear
        SQLQ = SQLQ & " AND PP_NBR >= " & txtWeek.Text & " AND PP_NBR <= " & txtWeekTo.Text
        If Len(clpPayP.Text) > 0 Then
            SQLQ = SQLQ & " AND PP_PAYP in ('" & Replace(clpPayP.Text, ",", "','") & "')"
        End If
        SQLQ = SQLQ & " ORDER BY PP_NBR"
        rsPayPrd.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        If Not rsPayPrd.EOF Then
            xCol = 3
            rsPayPrd.MoveFirst
            Do While Not rsPayPrd.EOF
                If glbCompSerial = "S/N - 2257W" Then   'Hamilton CAS - Ticket #15000
                    exSheet.Cells(4, xCol) = Format(rsPayPrd("PP_END"), "mmm dd, yyyy")
                Else
                    exSheet.Cells(4, xCol) = rsPayPrd("PP_YEAR") & "/" & rsPayPrd("PP_NBR")
                End If
                
                xCol = xCol + 1
                rsPayPrd.MoveNext
            Loop
        End If
        
        xRow = 5
        'Columns: 1 - Number/Name, 2 - Pay Period, 3 - Pay Period, 4 - Pay Period, etc. upto 12 Pay Periods
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            noStatus = True
            
            'Display Employee's Pay Period Status
            rsPayPrd.Requery
            If Not rsPayPrd.EOF Then
                rsPayPrd.MoveFirst
                xCol = 3
                Do While Not rsPayPrd.EOF
                    xStatus = get_TSStatus(rsHREmp("ED_EMPNBR"), rsPayPrd("PP_START"), rsPayPrd("PP_END"), rsPayPrd("PP_NBR"))
                    
                    'Check if Timesheet Status matches user selection criteria
                    Select Case xStatus
                        Case chkNotEntered.Tag
                            If chkNotEntered Then noStatus = False Else xStatus = ""
                        Case chkSaved.Tag
                            If chkSaved Then noStatus = False Else xStatus = ""
                        Case chkApproved.Tag
                            If chkApproved Then noStatus = False Else xStatus = ""
                        Case chkSubmitted.Tag
                            If chkSubmitted Then noStatus = False Else xStatus = ""
                        Case chkReSubmitted.Tag
                            If chkReSubmitted Then noStatus = False Else xStatus = ""
                        Case chkRejected.Tag
                            If chkRejected Then noStatus = False Else xStatus = ""
                        Case chkAppFwd.Tag 'Ticket #27551 Franks 09/17/2015
                            If chkAppFwd Then noStatus = False Else xStatus = ""
                        Case "Leave of Absence"
                            noStatus = False
                    End Select
                    If noStatus Then xStatus = ""
                    
                    If xStatus <> "" Then
                        If xStatus = "APP/FWD" Then 'Ticket #27551 Franks 09/17/2015
                            exSheet.Cells(xRow, xCol) = "Approve/Forward"
                        Else
                            exSheet.Cells(xRow, xCol) = Left(UCase(xStatus), 1) & Right(LCase(xStatus), Len(xStatus) - 1)
                        End If
                    End If
                    
                    xCol = xCol + 1
                    rsPayPrd.MoveNext
                Loop
            End If
            
            'Skip this employee if Timesheet Status does not match user selection criteria
            If noStatus Then
                GoTo Next_Employee
            End If
            'Display Employee # and Employee Name
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            xEmpDisp = xEmpDisp + 1
            
            xRow = xRow + 1
Next_Employee:
            rsHREmp.MoveNext
        Loop
        rsPayPrd.Close
                
        exSheet.Cells(xRow + 2, 2) = "Total Number of Employees : " & xEmpDisp
        exSheet.Rows(xRow + 2).Font.Bold = True

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
            
Exit Sub

Export_Timesheet_Status_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "Timesheet Status", "SELECT")
Resume Next
            
End Sub

Private Function get_TSStatus(xEmpnbr, strPPStartDate, strPPEndDate, xPPNbr)
Dim SQLQ, statusFlag
Dim rsDS As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection
Dim xEmpStatus As String
Dim xLOA As Boolean

    If glbSQL Or glbOracle Then
        Set gdbESS = gdbAdoIhr001
    Else
        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
    End If
    
    On Error Resume Next
    SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_TIMESHEET "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND AD_PPSTART >=" & Date_SQL(strPPStartDate)
    SQLQ = SQLQ & " AND AD_PPEND <=" & Date_SQL(strPPEndDate)
    SQLQ = SQLQ & " AND AD_PPNBR =" & xPPNbr

    If glbSQL Or glbOracle Then
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsDS.Open SQLQ, gdbESS, adOpenForwardOnly
    End If
    
    get_TSStatus = ""
    statusFlag = True
    
    If rsDS.EOF Then
        xEmpStatus = GetEmpData(xEmpnbr, "ED_EMP")
        xLOA = IsLOATypeCode(xEmpStatus)
        If xLOA Then
            get_TSStatus = "Leave of Absence"
        Else
            get_TSStatus = "Not Entered"
        End If
    Else
        Do While Not rsDS.EOF
            If statusFlag Then
                If IsNull(rsDS("AD_APPROVED")) Then
                    If rsDS("AD_UPLOAD") & "" = "Y" Then
                        get_TSStatus = "SUBMITTED"
                    Else
                        get_TSStatus = "SAVED"
                    End If
                Else
                    get_TSStatus = rsDS("AD_APPROVED")
                    
                    If get_TSStatus = "RESUBMIT" Then get_TSStatus = "RESUBMITTED"
                    
                End If
                statusFlag = False
            Else
                'getStatus="Inconsistent"
                get_TSStatus = "SAVED"
            End If
            rsDS.MoveNext
        Loop
        rsDS.Close
    End If
    
    If get_TSStatus = "" Then
        SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(strPPStartDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(strPPEndDate)
    
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        statusFlag = True
        If Not rsDS.EOF Then
            get_TSStatus = "APPROVED"
        End If
        rsDS.Close
    End If
    Set rsDS = Nothing
    
    If Err.Number <> 0 Then
    End If
    
End Function

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

'Private Function LOA_Type(xEmpStatus)
'    Dim rsHRTable As New ADODB.Recordset
'    Dim SQLQ As String
'
'    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_KEY = '" & xEmpStatus & "' AND TB_NAME = 'EDEM' "
'    rsHRTable.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    'xStr = ""
'    If Not rsHRTable.EOF Then
'        If rsHRTable("TB_USR3") = True Then
'            LOA_Type = True
'        Else
'            LOA_Type = False
'        End If
'    Else
'        LOA_Type = False
'    End If
'    rsHRTable.Close
'
'End Function
